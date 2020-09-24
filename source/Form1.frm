VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmMain 
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox border 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   720
      MousePointer    =   7  'Size N S
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox border 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      MousePointer    =   7  'Size N S
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuParent 
      Caption         =   "parent"
      Visible         =   0   'False
      Begin VB.Menu mnuChild 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuNewWindow 
      Caption         =   "newwindow"
      Visible         =   0   'False
   End
   Begin VB.Menu mnusrc 
      Caption         =   "src"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private thewn As CWordNetYasu
Const ccIncrement = 50000 '==========================================
Private wbctrl_top_min_height
Private wbctrl_bottom_min_height
Private Const default_wbctrl_top_height = 800
Private Const default_wbctrl_bottom_height = 300
Dim ccoffset As Long
Dim ccoffset2 As Long
Dim ccoffset3 As Long
Private wbctrl_top_height
Private wbctrl_bottom_height
Private Dragging As Boolean
Private oldy
Public curword
Private dict_loading As Boolean
Public local_mode As Boolean
Public external_resource_dir As String


Public curword_found As Boolean
Public main_document_complete As Boolean
Public indexdir As String

Dim WithEvents htmldoc1 As MSHTML.HTMLDocument
Attribute htmldoc1.VB_VarHelpID = -1
Dim WithEvents htmldoc2 As MSHTML.HTMLDocument
Attribute htmldoc2.VB_VarHelpID = -1
Dim WithEvents htmldoc3 As MSHTML.HTMLDocument
Attribute htmldoc3.VB_VarHelpID = -1
Dim WithEvents htmldoc4 As MSHTML.HTMLDocument
Attribute htmldoc4.VB_VarHelpID = -1
Dim WithEvents htmldoc5 As MSHTML.HTMLDocument
Attribute htmldoc5.VB_VarHelpID = -1
Dim WithEvents htmldoc6 As MSHTML.HTMLDocument
Attribute htmldoc6.VB_VarHelpID = -1
Dim WithEvents htmldoc7 As MSHTML.HTMLDocument
Attribute htmldoc7.VB_VarHelpID = -1
Dim WithEvents htmldoc8 As MSHTML.HTMLDocument
Attribute htmldoc8.VB_VarHelpID = -1
Dim WithEvents htmldoc9 As MSHTML.HTMLDocument
Attribute htmldoc9.VB_VarHelpID = -1


         Const ConcatStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


         Private Declare Function GetTickCount Lib "kernel32" () As Long







Private Sub ArrangeControls()
'Dim x As InternetExplorer
'On Error GoTo ErrorHandler

border_thickness = 100
infrm_top = 0
infrm_left = 0
infrm_width = Me.ScaleWidth
infrm_height = Me.ScaleHeight
If wbctrl_top_height < wbctrl_top_min_height Then
    wbctrl_top_height = wbctrl_top_min_height
End If
If wbctrl_bottom_height < wbctrl_bottom_min_height Then wbctrl_bottom_height = wbctrl_bottom_min_height

wbctrl_middle_height = infrm_height - wbctrl_top_height - wbctrl_bottom_height - 2 * border_thickness
'Debug.Print wbctrl_middle_height
If wbctrl_middle_height <= 0 Then
    'OOOPS  THERE'S NO SPACE FOR THE MAIN WINDOW
    wbctrl_middle_height = 0
    remaining_window_height = infrm_height - 2 * border_thickness
    If remaining_window_height <= 0 Then
        'DOOMED.  CAN'T DRAW ANYTHING
        Debug.Print "doomed"
        Exit Sub
    Else
    'OK.  DIVIDE UP THE REMAININGS
        wbctrl_top_height = remaining_window_height * wbctrl_top_height / (wbctrl_top_height + wbctrl_bottom_height)
        wbctrl_bottom_height = remaining_window_height - wbctrl_top_height
        Debug.Print "divided"
    End If
Else

End If
first_loop = True
For Each wbctrl In wb
    wbctrl.Width = infrm_width '- 2 * Screen.TwipsPerPixelX
    wbctrl.Left = infrm_left '+ 2 * Screen.TwipsPerPixelX
Next
For Each bdctrl In border
    bdctrl.Width = infrm_width
    bdctrl.Left = infrm_left
    bdctrl.Height = border_thickness
Next

wb(0).Top = infrm_top
wb(0).Height = wbctrl_top_height
nextwindowpos = wb(0).Top + wb(0).Height

border(0).Top = nextwindowpos

If wbctrl_middle_height > 0 Then
    'wb(1).Visible = True
    wb(1).Top = nextwindowpos + border_thickness
    wb(1).Height = wbctrl_middle_height
    nextwindowpos = wb(1).Top + wb(1).Height
Else
    'wb(1).Visible = True

    nextwindowpos = nextwindowpos + border_thickness
        wb(1).Top = nextwindowpos
End If

border(1).Top = nextwindowpos

wb(2).Top = nextwindowpos + border_thickness
wb(2).Height = wbctrl_bottom_height


ErrorHandler:

End Sub

Private Sub border_DblClick(index As Integer)
Select Case index
Case 0
    If wbctrl_top_height = 0 Then
        wbctrl_top_height = default_wbctrl_top_height
    Else
        wbctrl_top_height = 0
    End If
    
Case 1
    If wbctrl_bottom_height = 0 Then
        wbctrl_bottom_height = default_wbctrl_bottom_height
    Else
        wbctrl_bottom_height = 0
    End If
End Select
ArrangeControls
End Sub

Private Sub border_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dragging = True
    

oldy = 0
End Sub

Private Sub border_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 
    If Not Dragging Then Exit Sub
      
    If oldy = 0 Then
        'Debug.Print "========"
        oldy = y

    Else

        If index = 0 Then
            
                wbctrl_top_height = y - oldy + wbctrl_top_height
                'Debug.Print wbctrl_top_height
        ElseIf index = 1 Then
            wbctrl_bottom_height = wbctrl_bottom_height - (y - oldy)
        End If
    End If
     


ArrangeControls
    
End Sub

Private Sub border_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dragging = False


    

End Sub

Private Sub Command1_Click()
    For i = 1000 To 3000
        wbctrl_top_height = i
        ArrangeControls
        
    Next i
End Sub

Private Sub Form_Load()
 
external_resource_dir = "http://www.yasuaki.com/eetp/external/"
local_mode = False
indexdir = "C:\WN16\DICT"
main_document_complete = False
curword_found = False
wb(0).Navigate2 external_resource_dir & "html/navigation.htm"
wb(1).Navigate2 ("about:blank")
'wb(2).Navigate2 ("about:blank")
wb(2).Navigate2 external_resource_dir & "html/blankframe.htm"


wbctrl_top_height = default_wbctrl_top_height
wbctrl_bottom_height = default_wbctrl_bottom_height
If local_mode = True Then
    Set thewn = New CWordNetYasu
    thewn.Construct (indexdir)
End If
'hHook = SetWindowsHookEx(2, AddressOf Keyboard, App.hInstance, 0)
ChrTrap = vbKeyF10
'wb(0).Navigate2 "http://microsoft.com"

Debug.Print App.Path + "\webnavigation\index.htm"
Dim MyFile, MyPath, MyName



ArrangeControls
End Sub
Public Function statustext(txt As String)
On Error Resume Next
'wb(2).Document.frames(0).Document.body.innerText = txt
End Function


Private Sub Form_Unload(Cancel As Integer)
If local_mode Then
    thewn.Destruct
End If


End Sub
Private Sub Form_Resize()
ArrangeControls
End Sub



Private Function htmldoc1_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc2_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc3_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc4_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc5_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc6_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc7_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc8_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc9_oncontextmenu() As Boolean
    htmldoc_oncontextmenu
End Function
Private Function htmldoc_oncontextmenu() As Boolean
    Debug.Print "IDocHostUIHandler_ShowContextMenu"
    If Me.main_document_complete = True Then
        Dim NewPopup As New frmPopup
        Dim XY As POINTAPI

        GetCursorPos XY

        NewPopup.PopupNow Me, XY.x, XY.y
    Else
        statustext "document not finished loading"
    End If
End Function

Private Sub mnuNewWindow_Click()
On Error Resume Next
Dim frmWB As frmMain
Set frmWB = New frmMain
frmWB.wb(1).RegisterAsBrowser = True

frmWB.Visible = True
End Sub

Private Sub mnusrc_Click()
    wb(2).Document.frames(0).Document.body.innerText = wb(1).Document.body.outerHTML
End Sub

Private Sub wb_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)


If index <> 1 Then Exit Sub
main_document_complete = True
If wb(1).Document.frames.length = 0 Then

Set htmldoc1 = wb(1).Document

RecurseFrames wb(1).Document

Else
On Error Resume Next
    Set htmldoc1 = wb(1).Document.frames(0).Document
    RecurseFrames htmldoc1
    Set htmldoc2 = wb(1).Document.frames(1).Document
    RecurseFrames htmldoc2
    Set htmldoc3 = wb(1).Document.frames(2).Document
    RecurseFrames htmldoc3
    Set htmldoc4 = wb(1).Document.frames(3).Document
    Set htmldoc5 = wb(1).Document.frames(4).Document
    Set htmldoc6 = wb(1).Document.frames(5).Document
    Set htmldoc7 = wb(1).Document.frames(6).Document
    Set htmldoc8 = wb(1).Document.frames(7).Document
    Set htmldoc9 = wb(1).Document.frames(8).Document
On Error GoTo 0
End If
 'wb(2).Document.frames(0).Document.body.innerText = r
End Sub
Private Sub RecurseFrames(ByVal iDoc As HTMLDocument)
    Dim Frame As HTMLFrameElement
    Dim Range As IHTMLTxtRange
    Dim Title As String
    Dim TextInfo As String

    
    On Error Resume Next
    
    Title = iDoc.Title
    If Title = "" Then
        Title = iDoc.parentWindow.Name
    End If
    Dim i As Long

    If inode Is Nothing Then
        'if this is the first time, add a root node
        Set tvNode = tvTreeView.Nodes.Add(, , , Title)
        tvNode.Expanded = True
        'Set iNode = tvTreeView.Nodes.Add(, , , Title)
        'Debug.Print iNode

    Else
        Set tvNode = tvTreeView.Nodes.Add(inode.index, tvwChild, , Title)
    End If
    
    TextInfo = "Title: " & Title & " {" + vbCrLf
    
    'check to see if the document has a BODY
    If iDoc.body.tagName = "BODY" Then
        'fill the tree with following collections

        
        
        FillTree2 iDoc, "OBJECT"
        'use the text range object to get text out of BODY
        Set Range = iDoc.body.createTextRange
        TextInfo = TextInfo & Range.Text & vbCrLf
        Set Range = Nothing
    End If
    
    txtText.Text = txtText.Text & TextInfo & "}" & vbCrLf
    
    'recurse all the frames
    For Each Frame In iDoc.frames
        RecurseFrames Frame.Document
    Next
End Sub
Private Sub FillTree2(iDoc As HTMLDocument, iMatchTag As String)
    Dim Element As Object
    Dim Info As String

    
    Dim testa As HTMLBody
    On Error Resume Next
             Dim c As New Collection
            c.Add "beforeBegin"
            c.Add "afterBegin"
            c.Add "beforeEnd"
            c.Add "afterEnd"

    For Each Element In iDoc.All
                Set testa = Element

                 Dim i As Long
                Dim s As String
    If False Then
        thetagname = testa.tagName
        r = InStr(1, LCase(thetagname), "body")
        
    
        If r = Null Then r = 0
        If r <> 0 Then
            'testa.innerHTML =  & testa.innerHTML
        End If
    End If
       ' var nod=document.createElement("B");
       ' document.body.insertBefore(nod);
       ' nod.innerText = "A New bold object has been"
       '             inserted into the document."
    '}
              

     
     
     
     
                             Dim testc As HTMLBody
                        Set testc = Element
                   '     Debug.Print testa.toString
                  '     Debug.Print testc.toString
                       
                       If testc.hasChildNodes = True Then
                      'Debug.Print testc.childNodes.length
                       End If
                       
                For i = 1 To 3 Step 2
                    
                    
                    s = ""
                     s = testa.getAdjacentText(c.Item(i))
                    's2 = testa.getAdjacentText(c.Item(i))
                     testa.insertBefore (nod)
                    '                      nod.innerText = "asdf"
                    ' testa.Document.body.insertBefore (nod)

                    'If s <> Empty Then testa.insertAdjacentText "x.item(i)", "("
                    'Debug.Print testa.outerHTML
                    If s <> Empty And Len(s) > 1 Then
                        'Element.innerHTML = "<b>" & (Element.innerHTML) & "</b>"
                        'testa.replaceAdjacentText c.Item(i), "[" + testa.getAdjacentText(c.Item(i)) + "]"
                        Dim NewText As String
                        NewText = spantext(testa.getAdjacentText(c.Item(i)))
                        'Debug.Print newtext
                        testa.replaceAdjacentText c.Item(i), ""
                       testa.insertAdjacentHTML c.Item(i), NewText


                        'wb(2).Document.body.innerText = wb(1).Document.body.outerHTML
                       'Exit For
                    End If

                Next
                    
        Next
End Sub
Private Sub wb_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)

On Error Resume Next
Dim frmWB As frmMain
Set frmWB = New frmMain
frmWB.wb(1).RegisterAsBrowser = True
Set ppDisp = frmWB.wb(1).object
frmWB.Visible = True


End Sub



Private Function navcmd(cmdtype, cmdstr)
On Error Resume Next

Select Case cmdtype
Case "gourl"
    'Debug.Print "url"
    wb(1).Navigate2 (cmdstr)
Case "cmd"
    'Debug.Print "cmd"
    Select Case cmdstr
    Case "goforward"
        wb(1).GoForward
    Case "goback"
        wb(1).GoBack
    Case "home"
        wb(1).Navigate ("http://www.yasuaki.com/eetp")
    End Select
Case "cw"
    'Debug.Print cmdstr
        curword_found = False
        curword = LCase(cmdstr)
    
        Me.Caption = "The cursor is over " + curword

        Dim curword2 As String
        curword2 = curword
        If local_mode = True Then
            swrite = ""
            swrite = thewn.GetXMLIdx(curword2)
        
            If swrite <> "" Then
                curword_found = True
                fn = FreeFile
                Open external_resource_dir & "xml\idx.xml" For Output As fn
                    swrite = thewn.GetXMLIdx(curword2)
                    Print #fn, swrite
                Close fn
                'wb(2).Navigate App.Path & "\external\xml\idx.xml"
            End If

        End If
End Select

End Function
Private Sub RecurseFrames2(ByVal iDoc As HTMLDocument)
    Dim Frame As HTMLFrameElement
    Dim Range As IHTMLTxtRange
    Dim Title As String
    Dim TextInfo As String

    
    On Error Resume Next
    
    Title = iDoc.Title
    If Title = "" Then
        Title = iDoc.parentWindow.Name
    End If
    Dim i As Long

    If inode Is Nothing Then
        'if this is the first time, add a root node
        Set tvNode = tvTreeView.Nodes.Add(, , , Title)
        tvNode.Expanded = True
        'Set iNode = tvTreeView.Nodes.Add(, , , Title)
        'Debug.Print iNode

    Else
        Set tvNode = tvTreeView.Nodes.Add(inode.index, tvwChild, , Title)
    End If
    
    TextInfo = "Title: " & Title & " {" + vbCrLf
    
    'check to see if the document has a BODY
    If iDoc.body.tagName = "BODY" Then
        'fill the tree with following collections

        
        
        FillTree2 iDoc, "OBJECT"
        'use the text range object to get text out of BODY
        Set Range = iDoc.body.createTextRange
        TextInfo = TextInfo & Range.Text & vbCrLf
        Set Range = Nothing
    End If
    
    txtText.Text = txtText.Text & TextInfo & "}" & vbCrLf
    
    'recurse all the frames
    For Each Frame In iDoc.frames
        RecurseFrames Frame.Document
    Next
End Sub

Private Sub walktags(otags, r)
    'Dim otag As MSHTML.HTMLHtmlElement
    On Error GoTo 0
    Dim otag As HTMLHtmlElement
    On Error Resume Next
    For Each otag In otags
        

        If otag.childNodes.length > 1 Then
            
            
            walktags otag.childNodes, r

        Else
        Debug.Print otag.childNodes.length, otag.tagName, otag.innerText
            t = LCase(otag.tagName)
            If LCase(otag.tagName) = "a" Then r = r & "[" & otag.tagName & ":" & otag.innerText & "]"
            If t = "a" Then
                'Debug.Print otag.innerText
                'otag.innerText = UCase(otag.innerText)
            End If
        End If
        
    Next

    
End Sub

Private Function spantext(sorg) As String
    Dim i As Long
    
    Dim prevwhite As Boolean
    Dim word As String

    Dim startpos As Long
    Dim whStartpos As Long
    Dim newword As String
        ccoffset = 0
    prevwhite = True
    For i = 1 To Len(sorg)
        thechar = Mid$(sorg, i, 1)
        isnormal = True
        On Error GoTo conthere
        
        a = Asc(thechar)
        If ((0 <= a And a < 65) Or (97 > a And a > 90) Or (255 >= a And a > 122)) Then
        
            isnormal = False
        End If
        
conthere:
        If isnormal = True Then
            If prevwhite = True Then
                startpos = i
            End If
            prevwhite = False
        Else
            
            If prevwhite = False Then
    
        
        newword = wordout(Mid$(sorg, startpos, i - startpos))

         word = word + newword
    
        '  word = Left$(word, ccoffset)
  
                
                'Debug.Print word
            End If

            'Concat word, "sdfa"
    
          'word = Left$(word, ccoffset)
            prevwhite = True
            word = word + thechar
        End If
            
    Next
   
            If prevwhite = False Then

        newword = wordout(Mid$(sorg, startpos, i - startpos))

         word = word + newword

            End If
'Debug.Print word
    spantext = word
    'Debug.Print word
End Function
Private Sub wb_StatusTextChange(index As Integer, ByVal Text As String)

    colonstr = "_ee://"
    colonpos = InStr(1, Text, colonstr, vbTextCompare)
    colonlen = Len(colonstr)
    If colonpos <> 0 Then
        cmdtype = Left(Text, colonpos - 1)
        cmdstr = Mid(Text, colonpos + colonlen)
        navcmd cmdtype, cmdstr
    Else
        If index = 1 Then statustext (Text)
    End If

End Sub
Private Function wordout(word As String) As String

        
        wordout = wordout + "<SPAN id=cw onmouseout=""window.status='cw_ee://';"" onmouseover =""window.status='cw_ee://"

          

        wordout = wordout + word
    

        wordout = wordout + "';"">"

    
       wordout = wordout + word
 
          
         wordout = wordout + "</SPAN>"

End Function

Private Sub wb_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'frmMain.Caption = URL & "processing text.  please wait..."
main_document_complete = False
On Error Resume Next
If index = 1 Then
    For i = 0 To wb(0).Document.frames(0).Document.All.length - 1
        id_name = wb(0).Document.frames(0).Document.All(i).Id
        If id_name = "txturl" Then
            Dim otag As MSHTML.HTMLInputElement
            Set otag = wb(0).Document.frames(0).Document.All(i)
            otag.Value = URL
        End If
    Next
End If
'sw(0).SetVariable "txturl", URL
End Sub

Private Sub sw_FSCommand(index As Integer, ByVal command As String, ByVal args As String)
    navcmd command, args
End Sub

