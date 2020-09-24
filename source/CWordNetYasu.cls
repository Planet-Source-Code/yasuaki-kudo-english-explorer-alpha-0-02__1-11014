VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWordNetYasu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private indexfile

Private Const total_wn_files = 10
Private wnf(1 To total_wn_files) As CWordNetFile
Private m_idx_type_names As Collection
Private postagnames(1 To 4) As String

'Private m_pos_type_ids As Collection
'Private Const total_pos_type_ids = 5



'Private idxheadloc As Collection

Public Function Construct(indexdir As String)


    For i = 1 To total_wn_files
        Set wnf(i) = New CWordNetFile
    Next i

    wnf(1).Construct indexdir, "NOUN.IDX"
    'm_pos_type_ids.Add 1, "n"
    wnf(2).Construct indexdir, "Verb.IDX"
    'm_pos_type_ids.Add 2, "v"
    wnf(3).Construct indexdir, "ADJ.IDX"
    'm_pos_type_ids.Add 3, "a"
    wnf(4).Construct indexdir, "ADV.IDX"
    'm_pos_type_ids.Add 4, "r"
    wnf(5).Construct indexdir, "SENSE.IDX"
    wnf(6).Construct indexdir, "GLOSS.IDX"
    wnf(7).Construct indexdir, "NOUN.DAT"
    wnf(8).Construct indexdir, "Verb.DAT"
    wnf(9).Construct indexdir, "ADJ.DAT"
    wnf(10).Construct indexdir, "ADV.DAT"
    'm_pos_type_ids.Add 4, "r"
    Set m_idx_type_names = New Collection
    m_idx_type_names.Add "lemma"
    m_idx_type_names.Add "pos"
    m_idx_type_names.Add "poly_cnt"
    m_idx_type_names.Add "p_cnt"
    m_idx_type_names.Add "ptr_symbol"
    m_idx_type_names.Add "sense_cnt"
    m_idx_type_names.Add "tagsense_cnt"
    m_idx_type_names.Add "synset_offset"
    postagnames(1) = "n"
    postagnames(2) = "v"
    postagnames(3) = "a"
    postagnames(4) = "r"
End Function
Public Function Destruct()
    For i = 1 To total_wn_files
        wnf(i).Destruct
    Next i
End Function



Public Function GetXMLIdx(sword As String) As String
    something_found = False
    For i = 1 To 4
        sposline = getxmlpos(sword, i)
        If sposline <> Empty Then
            something_found = True
            'sresult = sresult + xmlout(sposline, postagnames(i))
            sresult = sresult & xmlout(xmlout(postagnames(i), "t") & sposline, "p")
        End If
        
    Next i
    If something_found = True Then
    
        sresult = xmlout(UCase(sword), "k") + sresult
        header = "<?xml version='1.0'?><?xml-stylesheet type=""text/xsl"" href=""idx.xsl"" ?>"
        GetXMLIdx = header + xmlout(sresult, "idx")
    End If
    
End Function
Private Function getxmlpos(sword As String, posno) As String
    getxmlpos = Empty
    If wnf(posno).SeekByWord(sword, line) <> 0 Then
        Exit Function
    End If
    elms = StrSplit(line, " ")
    howmany = elms(2)
    fromwhere = 6 + elms(3)
    Dim i As Long
    For i = fromwhere To fromwhere + howmany - 1
         'sresult = sresult & elms(i) & vbCrLf
         sresult = sresult + xmlout(getxmlsynset(Int(elms(i) + 1), posno), "s")
    Next
    getxmlpos = sresult

End Function
Private Function getxmlsynset(seekloc As Long, posno) As String
    getxmlsynset = Empty
    Dim elms, elms2
    wnf(posno + 6).Seek2 (seekloc)
    line = wnf(posno + 6).ReadLine
    If line = Empty Then
        Exit Function
    End If
    elms = StrSplit(line, "|")
    elms2 = StrSplit(elms(0), " ")
    'wcnt = elms2(3)
    wcnt = Val("&H" & elms2(3))
    For i = 0 To wcnt - 1
        wd = elms2(4 + i * 2)
        If wd <> key Then
            wd = Replace(wd, "_", " ")
         sresult = sresult + xmlout(wd, "w")
        End If
    
    Next
    p_cnt = Int(elms2(4 + 2 * wcnt))
    gloss = xmlproperstring(elms(1))
    sresult = sresult + xmlout(gloss, "g")
    getxmlsynset = sresult

End Function

Private Function xmlproperstring(line)
    line = Replace(line, "&", "&amp;")
    line = Replace(line, "<", "&lt;")
    line = Replace(line, ">", "&gt;")
    xmlproperstring = line
End Function
Private Function xmlout(line, Tag, Optional optarg = "")
    sresult = "<" + Tag
    If optarg <> "" Then
        sresult = sresult + " "
    End If
    sresult = sresult + optarg + ">"
    sresult = sresult + line
    sresult = sresult + "</" + Tag + ">"
    xmlout = sresult
End Function

Private Function xmloutsynset(line, key):
    Dim elms
    elms = StrSplit(line, "|")
    elms = StrSplit(elms(0), " ")
    sresult = "\n"

End Function
