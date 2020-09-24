VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmPopup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5250
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   495
      Left            =   4680
      OleObjectBlob   =   "Popup.frx":0000
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser doodad 
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      ExtentX         =   7011
      ExtentY         =   3625
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
   Begin VB.Timer FocusChecker 
      Interval        =   50
      Left            =   855
      Top             =   765
   End
   Begin VB.Menu mnuRead 
      Caption         =   "read it!"
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Windows API Declarations
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Integer

Private MyWindow As Integer
Private WithEvents MyParentForm As Form
Attribute MyParentForm.VB_VarHelpID = -1

Public Sub PopupNow(ParentForm As Form, vleft, vtop)
    Dim NewPopup As New frmPopup, Point As POINTAPI
    Set MyParentForm = ParentForm
    
    'ClientToScreen ParentControl.hwnd, Point
    'Me.Left = Point.x * Screen.TwipsPerPixelX
    'Me.Top = Point.y * Screen.TwipsPerPixelY + ParentControl.Height
    'Me.Width = ParentControl.Width
    Dim x As Long, y As Long
    x = vleft * Screen.TwipsPerPixelX
    y = vtop * Screen.TwipsPerPixelY
    If x + Me.Width > Screen.Width Then
        x = x - Me.Width
    End If
    If y + Me.Height > Screen.Height Then
        y = y - Me.Height
    End If
    Me.Left = x
    Me.Top = y
    Me.Show
    Me.Caption = UCase(MyParentForm.curword)
    MyWindow = GetActiveWindow
    If MyParentForm.local_mode = True Then
    
        If MyParentForm.curword_found = True Then
            doodad.Navigate2 App.Path & "\external\xml\idx.xml"
        Else
            doodad.Navigate2 App.Path & "\external\html\curwordnotfound.htm"
        End If
    Else
            doodad.Navigate2 "http://www.yasuaki.com/eetp/external/eewnburst.cgi?word=" & MyParentForm.curword
    End If
    'doodad.Navigate2 "http://www.microsoft.com"
End Sub

Public Sub Done()
    Unload Me
End Sub



Private Sub FocusChecker_Timer()
    If GetActiveWindow <> MyWindow Then
        Done
    End If
End Sub

Private Sub Form_Resize()
    doodad.Top = 0
    doodad.Left = 0
    doodad.Width = Me.ScaleWidth
    doodad.Height = Me.ScaleHeight
End Sub

Private Sub mnuRead_Click()
    On Error Resume Next
    DirectSS1.Speak doodad.Document.body.innerText
    
End Sub

Private Sub MyParentForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Done
End Sub
