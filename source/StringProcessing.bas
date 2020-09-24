Attribute VB_Name = "StringProcessing"
'****************************************************************
' String Processing Library
'
' Created 3 May 2000 by James Vincent Carnicelli
' Updated 1 June 2000 by James Vincent Carnicelli
'   http://alexandria.nu/user/jcarnicelli/
'
' Notes:
' These routines exist to simplify basic string processing,
' including parsing, translation, and validation.
'****************************************************************

Option Explicit

Public Const CHARSET_LETTERS_UCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const CHARSET_LETTERS_LCASE = "abcdefghijklmnopqrstuvwxyz"
Public Const CHARSET_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Public Const CHARSET_DIGITS = "0123456789"
Public Const CHARSET_ALPHANUMERIC = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
Public Const CHARSET_PRINTABLE = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
Public Const CHARSET_WHITESPACE = vbCr & vbLf & vbTab & " "

'Replace every occurrance of SearchFor with ReplaceWith.
'The result is returned, but Text is not modified.
'Avoid using vbTextCompare, which can double the execution
'time.  Same functionality as VB's new Replace() function.
Public Function StrReplace(ByVal Text As String, ByVal SearchFor As String, ByVal ReplaceWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim Pos As Long, NewText As String, SearchForLength As Long
    SearchForLength = Len(SearchFor)
    Do
        Pos = InStr(1, Text, SearchFor, Compare)
        If Pos = 0 Then Exit Do
        NewText = NewText & Left(Text, Pos - 1) & ReplaceWith
        Text = Mid(Text, Pos + SearchForLength)
    Loop
    NewText = NewText & Text
    StrReplace = NewText
End Function

'Split a string up into an array, splitting wherever SplitOn
'is found.  Same functionality as VB's new Split() function.
Public Function StrSplit(ByVal Text As String, SplitOn As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim List As Variant, Pos As Long, SplitOnLength As Long
    Dim Count As Integer
    List = Array()
    SplitOnLength = Len(SplitOn)
    Do
        Pos = InStr(1, Text, SplitOn, Compare)
        If Pos = 0 Then Exit Do
        ReDim Preserve List(Count)
        List(Count) = Left(Text, Pos - 1)
        Text = Mid(Text, Pos + SplitOnLength)
        Count = Count + 1
    Loop
    ReDim Preserve List(UBound(List) + 1)
    List(Count) = Text
    StrSplit = List
End Function

'Convert the hexadecimal number into a Long integer.  Hex string
'may not include formatting characters like '&' or 'H'.  This
'function compliments VB's bulit-in Hex() function.
Public Function FromHex(ByVal HexNumber As String) As Long
    Dim lPos As Long, sChar As String, nFactor As Integer
    Dim nLen As Integer
    nLen = Len(HexNumber)
    HexNumber = UCase(HexNumber)
    For lPos = 1 To nLen
        sChar = Mid(HexNumber, lPos, 1)
        Select Case sChar
            Case "F": nFactor = 15
            Case "E": nFactor = 14
            Case "D": nFactor = 13
            Case "C": nFactor = 12
            Case "B": nFactor = 11
            Case "A": nFactor = 10
            Case Else
                nFactor = CInt(sChar)
        End Select
        FromHex = FromHex + nFactor * 16 ^ (nLen - lPos)
    Next
End Function

'Pad the left side of the string with enough PadCharacter characters
'until the string is Size characters long.
Public Function LeftPad(Value, Size As Long, Optional PadCharacter As String = " ") As String
    LeftPad = "" & Value
    While Len(LeftPad) < Size
        LeftPad = PadCharacter & LeftPad
    Wend
End Function

'Pad the right side of the string with enough PadCharacter characters
'until the string is Size characters long.
Public Function RightPad(Value, Size As Long, Optional PadCharacter As String = " ") As String
    RightPad = "" & Value
    While Len(RightPad) < Size
        RightPad = RightPad & PadCharacter
    Wend
End Function

'Are all of the characters in Text letters (i.e., A - Z and a - z)?
Public Function AreLetters(Text As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(Text)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(Text, i, 1)))
        If nChar < 65 Or nChar > 90 Then Exit Function
    Next
    AreLetters = True
End Function

'Are all of the characters in Text digits (i.e., 0 - 9)?
Public Function AreDigits(Text As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(Text)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(Text, i, 1)))
        If nChar < 48 Or nChar > 57 Then Exit Function
    Next
    AreDigits = True
End Function

'Are all of the characters in Text letters or digits?
Public Function AreLettersOrDigits(Text As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(Text)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(Text, i, 1)))
        If (nChar < 65 Or nChar > 90) And (nChar < 48 Or nChar > 57) Then Exit Function
    Next
    AreLettersOrDigits = True
End Function

'Are all of the characters in Text within the specified set?
'An example of a set is "0123456789 .,Ee-", which should suffice to
'recognize most simple number formats and even scientific notation.
Public Function AreInSet(Text As String, CharacterSet As String, Optional CaseSensitive As Boolean = True) As Boolean
    Dim i As Long, lLen As Long
    lLen = Len(Text)
    For i = 1 To Len(Text)
        If InStr(1, CharacterSet, Mid(Text, i, 1), CaseSensitive + 1) = 0 Then Exit Function
    Next
    AreInSet = True
End Function

'Parse a string of values out of the given text stream.  A typical
'use would be to get information from an HTML file containing the
'following:
'
'  <H2> People: </H2>
'  <TABLE>
'    <TR><TH> Name </TH><TH> Social Security Number </TH></TR>
'    <TR><TD> John Doe </TD><TD> 123-45-6789 </TD></TR>
'    <TR><TD> Jane Doe </TD><TD> 987-65-4321 </TD></TR>
'  </TABLE>
'
'If you want to get a list of the people's names and SSNs, the
'following code would work (provided HtmlSource contains the
'contents of the HTML file):
'
'   Dim StartAt As Long, People, Person
'   People = Array(Array("Name", "SSN"))
'   StartAt = 1
'   While ParseN(HtmlSource, StartAt, _
'     Array( "<TR><TD>", "</TD><TD>", "</TD></TR>"), _
'     Person, StartAt)
'       ReDim Preserve People(UBound(People) + 1)
'       People(UBound(People)) = Person
'   Wend
'
'When it's done, the People array would contain a list of arrays
'that represent each person's name and SSN.  That is, its contents
'would be:
'
'   Name        SSN
'   John Doe    123-45-6789
'   Jane Doe    987-65-4321
'
'PosPastEnd contains the index of the first character right after
'the end of the last PatternArray pattern.  In this case, it would
'point to the new-line characters right after each successive
'"</TD></TR>" that follows the SSN.  PosPastEnd's value can, as the
'example here shows, be plugged into StartAt for the next iteration
'of ParseN().  ValueArray is a variant array you pass in that
'ParseN() fills with the list of values between each of the search
'patterns.  This will generally be what you're searching for, as
'well as any "random" garbage that might be in the middle (e.g.,
'HTML tag attributes you don't care about like "WIDTH=100").  If you
'want to keep the search for a segment from spilling over into some
'part of the text stream that you know marks the end of your
'legitimate data, set StopAt to point to the first character of the
'forbidden part of the data to parse.  This might be helpful if,
'say, you had three different HTML tables with the same kind of
'information and you wanted to process them separately.  You could
'search for "<TABLE" and "</TABLE>" in the beginning to find out
'where each table begins and ends, then use ParseN() with initial
'StartAt and StopAt attributes set to the beginning and end of each
'table chunk during the While loop.  Always be sure to check the
'return value to make sure it executed successfully.  ValueArray
'may be a misleading indicator, since it may be half-filled before
'ParseN() realizes the conditions are only partially met.
Public Function ParseN(Text As String, StartAt As Long, PatternArray As Variant, ValueArray As Variant, Optional PosPastEnd, Optional StopAt As Long = 0, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    Dim lStartPos As Long, lEndPos As Long
    Dim nValue As Integer, nPatternUbound As Integer

    If StopAt = 0 Then StopAt = Len(Text) + 1
    
    On Error Resume Next
    nPatternUbound = UBound(PatternArray)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "StringProcessing.ParseN()", "Expecting PatternArray to be an array; found instead a """ & TypeName(PatternArray) & """"
    End If
    If nPatternUbound < 1 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "StringProcessing.ParseN()", "PatternArray must contain at least two patterns to search for"
    End If
    
    If Not IsEmpty(ValueArray) Then
        nValue = UBound(ValueArray)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "StringProcessing.ParseN()", "Expecting ValueArray to be an array; found instead a """ & TypeName(ValueArray) & """"
        End If
    End If
    On Error GoTo 0
    
    ReDim ValueArray(nPatternUbound - 1)
    
    'PatternArray
    lStartPos = InStr(StartAt, Text, PatternArray(0))
    If lStartPos = 0 Then Exit Function
    lStartPos = lStartPos + Len(PatternArray(0))
    
    If lStartPos > StopAt Then Exit Function
    
    For nValue = 1 To nPatternUbound
        lEndPos = InStr(lStartPos, Text, PatternArray(nValue), Compare)
        If lEndPos = 0 Then Exit Function
        If lEndPos + Len(PatternArray(nValue)) > StopAt Then Exit Function
        ValueArray(nValue - 1) = Mid(Text, lStartPos, lEndPos - lStartPos)
        lStartPos = lEndPos + Len(PatternArray(nValue))
    Next
    
    PosPastEnd = lEndPos + Len(PatternArray(nPatternUbound))
    ParseN = True
End Function

