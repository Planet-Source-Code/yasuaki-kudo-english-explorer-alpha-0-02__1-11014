Attribute VB_Name = "Module3"
    Declare Function GetTempFileName Lib "kernel32" _
        Alias "GetTempFileNameA" _
        (ByVal lpszPath As String, _
        ByVal lpPrefixString As String, _
        ByVal wUnique As Long, _
        ByVal lpTempFileName As String) As Long
    
    Declare Function GetTempPath Lib "kernel32" _
        Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long

    Declare Function lstrlen Lib "kernel32" _
        Alias "lstrlenA" _
        (ByVal lpString As String) As Long


    Public Function DZBuildTempFileDir() As String
    '
    '  Local Variables
        Dim wrkTempFileName As String * 660
        Dim wrkTempFileDir As String * 660
        Dim wrkStringLength As Integer
        Dim wrkFlag As Long
        Dim wrkString As String
    '
    '  Set Error Handling
        On Error Resume Next
    '
    '  Set a Blank Name for the File
        wrkTempFileName = String(640, 0)
    '
    '  Assign a Default Name if Everything fails
     '   wrkString = JMRunPath("daisy432.tmp")
    '
    '  Set a Blank Name for the Temporary File Directory
        wrkTempFileDir = String(640, 0)
    '
    '  Get the Temporary File Directory
        wrkFlag = GetTempPath(600, wrkTempFileDir)
    '
    '  Get a Temporary File Name if Directory Found
        If (wrkFlag <> 0) Then
            wrkStringLength = lstrlen(wrkTempFileDir)
            wrkFlag = GetTempFileName(Left$(wrkTempFileDir, _
              wrkStringLength), "dsy", 0, wrkTempFileName)
            If (wrkFlag <> 0) Then
                'wrkStringLength = lstrlen(wrkTempFileName)
                wrkString = Left$(wrkTempFileName, wrkStringLength)
            End If
        End If
    '
    '  Return the Temporary File
        DZBuildTempFileDir = wrkString
    End Function


