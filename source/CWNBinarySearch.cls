VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWordNetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private m_filenumber As Integer
Private m_filepath As String
Private m_filesize As Long
Private m_filedir As String
Private m_filename As String
Private m_lastSeek As Long







Public Property Get filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.m_filename
    If IsObject(m_filename) Then
        filename = m_filename
    Else
        filename = m_filename
    End If
End Property

Public Function Seek2(seekloc As Long) As Long
    Seek2 = 1
    If seekloc < 1 Or seekloc > m_filesize Then
        Exit Function
    End If
    Seek #m_filenumber, seekloc
    m_lastSeek = Seek(m_filenumber)
    Seek2 = 0
End Function
Public Function ReadLine() As String
    ReadLineBySeek = Empty
    ReadLineBySeek = ""
    Dim sbuf As String
    sbuf = " "
    Seek #m_filenumber, m_lastSeek
    m_lastSeek = Seek(m_filenumber)
    Do
        Get #m_filenumber, , sbuf
        If EOF(m_filenumber) Or sbuf = Chr(10) Then
            Exit Do
        End If
        ReadLineBySeek = ReadLineBySeek & sbuf
    Loop
    ReadLine = ReadLineBySeek
End Function


'Public Function seekbyword(sm_filepath As String, ByVal key As String, sresults As Collection, Optional repeats As Integer = 1)
Public Function Construct(sm_filedir As String, sfilename As String)
    m_filedir = sm_filedir
    m_filename = sfilename
    m_filepath = m_filedir & "\" & m_filename
    m_filenumber = FreeFile
    m_filesize = FileLen(m_filepath)
    Open m_filepath For Binary Access Read As m_filenumber
    


End Function
Public Function Destruct()
    Close m_filenumber
End Function
Public Function SeekByWord(ByVal key As String, Optional line2 = "") As Integer
    SeekByWord = 1
    If key = "" Then
        Exit Function
    End If

    If Left(key, 1) = " " Then Exit Function
    
    key = key + " "
    keylen = Len(key)
    
    Dim start2  As Long
    Dim middle2  As Long
    Dim end2 As Long
    Dim offset2 As Long
    Dim line2isbig As Boolean

    start2 = 1
    end2 = m_filesize + 1

    currentdepth = 0
    Do
        If start2 >= end2 Then
            'offset2 = end2
            Exit Do
        End If
        
        middle2 = (start2 + end2) / 2

        line2 = nw(middle2, offset2, m_filenumber, m_filesize)
        'Debug.Print line2isbig & vbCrLf & start2 & " " & middle2 & " " & end2 & " " & line2;
        'If line2 = "" Then
        '    Exit Do
        'End If
        
        

        If line2 > key Or line2 = "" Then
            line2isbig = True
        Else
            line2isbig = False
        End If
        If offset2 > end2 Then
            If end2 = middle2 - 1 Then MsgBox ("endless loop")
            end2 = middle2 - 1
        ElseIf Left(line2, keylen) = key Then
            SeekByWord = 0
            Exit Do
        ElseIf line2isbig Then
            If end2 = middle2 - 1 Then MsgBox ("endless loop")
            end2 = middle2 - 1
        ElseIf line2isbig = False Then
            start2 = offset2 + Len(line2) - 1 ' PERFORMANCE TUNER ?? HOW EFFECTIVE?
        End If

        
    Loop

endproc:
    If Left(line2, keylen) = key Then SeekByWord = 0
        If Left(line2, 1) = " " Or Left(line2, Len(key)) < key Then
                middle2 = offset2
                line2 = nw(middle2, offset2, m_filenumber, m_filesize)
        End If
    m_lastSeek = offset2
End Function

Private Function nw(seekloc As Long, newseekloc As Long, m_filenumber As Integer, m_filesize As Long) As String 'nextword
    nw = ""
    
    Seek #m_filenumber, seekloc
    
    Dim sbuf As String
    sbuf = Space(1)
    
    stopchar = True
    sresult = ""
    stopchar_found = False
    lastloc = m_filesize
    Do
        Get #m_filenumber, , sbuf
         If EOF(m_filenumber) Then Exit Do
        If seekloc <= 1 Then
            seekloc = 1
            stopchar_found = True

        ElseIf seekloc >= lastloc Then
            seekloc = lastloc
        End If

        If stopchar_found Then
            If sbuf = Chr(10) Then
                
                Exit Do
            Else
                sresult = sresult & sbuf
            End If
        Else
            If sbuf = Chr(10) Then
                stopchar_found = True
                newseekloc = Seek(m_filenumber)
            End If
        End If
    Loop
    
    nw = sresult


End Function

