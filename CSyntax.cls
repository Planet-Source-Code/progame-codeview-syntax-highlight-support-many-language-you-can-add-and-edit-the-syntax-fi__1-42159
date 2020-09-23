VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'*ÀàÄ£¿éÃû³Æ£ºCSyntax
'*ÀàÄ£¿é¹¦ÄÜ£ºÓï·¨¸ßÁÁ
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-10-07  17:54:44
'***************************************************************
Option Explicit
Option Compare Binary

Public Event Progress(Value As Integer) '*´¦Àí½ø¶È

Private m_sTitle            As String   '*Óï·¨ÎÄ¼þTitle

Private m_sLnCommFlag1      As String   '*ÐÐ×¢ÊÍÍ·Ò»
Private m_sLnCommFlag2      As String   '*ÐÐ×¢ÊÍÍ·¶þ
Private m_sBlkCommHead1     As String   '*¿é×¢ÊÍÍ·Ò»
Private m_sBlkCommFoot1     As String   '*¿é×¢ÊÍÎ²Ò»
Private m_sBlkCommHead2     As String   '*¿é×¢ÊÍÍ·¶þ
Private m_sBlkCommFoot2     As String   '*¿é×¢ÊÍÎ²¶þ
Private m_objCommFont       As CFont    '*×¢ÊÍ×ÖÌåÉèÖÃ

Private m_sCommRtfHead      As String   '*×¢ÊÍ×ÖÌåRTF¸ñÊ½Í·
Private m_sCommRtfFoot      As String   '*×¢ÊÍ×ÖÌåRTF¸ñÊ½Î²

Private m_objNumFont        As CFont    '*Êý×Ö×ÖÌåÉèÖÃ

Private m_sNumRtfHead       As String   '*Êý×Ö×ÖÌåRTF¸ñÊ½Í·
Private m_sNumRtfFoot       As String   '*Êý×Ö×ÖÌåRTF¸ñÊ½Î²

Private m_iKwdCount         As Integer  '*¹Ø¼ü×ÖÀàÐÍÊý
Private m_sKwdSet()         As String   '*´æÈëÒ»Àà¹Ø¼ü×ÖµÄÈ«²¿×éºÏ(ÒÔchr(3)·Ö¸ô)
Private m_sKwdTitle()       As String   '*´Ë¹Ø¼ü×ÖÀàÐÍµÄÃû³Æ
Private m_objKwdFont()      As CFont    '*´Ë¹Ø¼ü×ÖÀàÐÍµÄ×ÖÌåÉèÖÃ
Private m_sKwdRtfHead()     As String   '*¹Ø¼ü×Ö×ÖÌåRTF¸ñÊ½Í·
Private m_sKwdRtfFoot()     As String   '*¹Ø¼ü×Ö×ÖÌåRTF¸ñÊ½Î²

Private m_bCase             As Boolean  '*ÊÇ·ñ´óÐ¡×ÖÃô¸Ð

Private m_sQuotFlag1        As String   '*×Ö·û´®±ß½çÒ»
Private m_sQuotFlag2        As String   '*×Ö·û´®±ß½ç¶þ
Private m_bMultiLine        As Boolean  '*ÊÇ·ñÐøÐÐÓÐÐ§
Private m_objQuotFont       As CFont    '*×Ö·û´®×ÖÌåÉèÖÃ
Private m_sQuotRtfHead      As String   '*×Ö·û´®×ÖÌåRTF¸ñÊ½Í·
Private m_sQuotRtfFoot      As String   '*×Ö·û´®×ÖÌåRTF¸ñÊ½Î²

Private m_sEscape           As String   '*ºóÃæ¸úÃô¸Ð×Ö·û£¬´Ë×Ö·ûºóÃæ×Ö·ûºöÂÔ

Private m_objColorManager   As CColor   '*ÑÕÉ«¹ÜÀíÀà
Private m_sColorRtf         As String   '*RTFÄÚÑÕÉ«ÐÅÏ¢

Private m_sWord()         As String     '*¶ÔÒ»ÐÐ×Ö·û´®·Ö¸ô³ÉÒ»¸ö´ÊÊý×é

Private m_iQuotState        As Integer  '*µ±Ç°ÊÇ·ñÔÚ×Ö·û´®ÄÚ[0..²»ÊÇ×Ö·û´®/1..ÔÚ×Ö·û´®Ò»ÄÚ/2..ÔÚ×Ö·û´®¶þÄÚ]
Private m_iCommState        As Integer  '*µ±Ç°×¢ÊÍ×´Ì¬[0..ÎÞ×¢ÊÍ/1..µ¥ÐÐ×¢ÊÍÒ»/2..µ¥ÐÐ×¢ÊÍ¶þ/3..¶àÐÐ×¢ÊÍÒ»/4..¶àÐÐ×¢ÊÍ¶þ]


Private Const DELIMITER = ",(){}[]-+*%/='~!&|\<>?:;."


Public Function ReadFile(filename As String) As Boolean
'*½«Óï·¨ÎÄ¼þ¶ÁÈë
Dim fn      As Integer
Dim s       As String
Dim i       As Integer
    fn = FreeFile()
    
    Open filename For Input As #fn
        Do While Not EOF(fn)
            Line Input #fn, s
            Call PreFormat(s)
            Call ReadDetail(fn, s)
        Loop
    Close #fn
    
    '*¶Ô×ÖÌåµÄ×Ö·û´®Ô¤ÏÈÉú³É
    m_sQuotRtfHead = m_objQuotFont.getRtfHead(m_objColorManager)
    m_sQuotRtfFoot = m_objQuotFont.getRtfTail
    m_sCommRtfHead = m_objCommFont.getRtfHead(m_objColorManager)
    m_sCommRtfFoot = m_objCommFont.getRtfTail
    
    m_sNumRtfHead = m_objNumFont.getRtfHead(m_objColorManager)
    m_sNumRtfFoot = m_objNumFont.getRtfTail
    
    m_sColorRtf = m_objColorManager.ComStr
    
    For i = 1 To m_iKwdCount
        m_sKwdRtfHead(i) = m_objKwdFont(i).getRtfHead(m_objColorManager)
        m_sKwdRtfFoot(i) = m_objKwdFont(i).getRtfTail
    Next i
    
    ReadFile = True

    Exit Function
    
Err_Proc:
    MsgBox Err.Description
    ReadFile = False
End Function

Private Sub PreFormat(ByRef s As String)
'*½øÓï·¨ÎÄ¼þÔ¤´¦Àí£¬ÒòÎªµ½Ê±ºò´úÂë»áÏÈ½øÐÐ¸ñÊ½»¯ÒÔÊÊÓ¦"}" "{" "\"
    s = Replace(Replace(Replace(s, "\", "\\"), "{", "\{"), "}", "\}")
End Sub

Private Sub ReadDetail(filenum As Integer, s As String)
'*¶ÁÈ¡Óï·¨ÎÄ¼þµÄ¾ßÌåÊµÏÖ
Dim sTmp    As String
Dim i       As Integer
Dim lstr    As String
Dim rstr    As String

    s = Trim(s)
    If Len(s) = 0 Then              'blank
        Exit Sub
    End If
    If Left(s, 1) = ";" Then        'comment
        Exit Sub
    End If
    
    '*¶ÁÈ¡×Ö·û´®µÄ×ÖÌåÅäÖÃ
    If Len(s) = 10 And Left(s, 10) = "#QUOTATION" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '¶ÁÈ¡ÏêÏ¸ÄÚÈÝ
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objQuotFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*¶ÁÈ¡×¢ÊÍµÄ×ÖÌåÅäÖÃ
    If Len(s) = 8 And UCase(Left(s, 8)) = "#COMMENT" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '¶ÁÈ¡ÏêÏ¸ÄÚÈÝ
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objCommFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*¶ÁÈ¡Êý×ÖµÄ×ÖÌåÅäÖÃ
    If Len(s) = 7 And UCase(Left(s, 7)) = "#NUMBER" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '¶ÁÈ¡ÏêÏ¸ÄÚÈÝ
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objNumFont)
            End If
        Loop
        
        Exit Sub
    End If
    
    '*¶ÁÈ¡¹Ø¼ü×Ö
    If UCase(Left(s, 8)) = "#KEYWORD" Then  'keyword head
    
        sTmp = Right(s, Len(s) - 9)
        m_iKwdCount = m_iKwdCount + 1
        
        ReDim Preserve m_sKwdSet(1 To m_iKwdCount)
        ReDim Preserve m_sKwdTitle(1 To m_iKwdCount)
        ReDim Preserve m_objKwdFont(1 To m_iKwdCount)
        ReDim Preserve m_sKwdRtfHead(1 To m_iKwdCount)
        ReDim Preserve m_sKwdRtfFoot(1 To m_iKwdCount)
        
        Set m_objKwdFont(m_iKwdCount) = New CFont
        
        m_sKwdTitle(m_iKwdCount) = sTmp
        m_sKwdSet(m_iKwdCount) = Chr(3)
        
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If

            '¶ÁÈ¡¹Ø¼ü×ÖµÄÏêÏ¸ÄÚÈÝ
            i = InStr(sTmp, "=")
            If i = 0 Or (i = 1 And Len(sTmp) = i) Then
                m_sKwdSet(m_iKwdCount) = m_sKwdSet(m_iKwdCount) & IIf(m_bCase, Trim(sTmp), UCase(Trim(sTmp))) & Chr(3)
            Else
                Call ReadFont(sTmp, i, m_objKwdFont(m_iKwdCount))
            End If
        Loop
        
        Exit Sub
    End If
    
    i = InStr(s, "=")
    If i = 0 Or i = 1 Or Len(s) = i Then
        Exit Sub
    End If
    
    '*¶ÁÈ¡»ù±¾ÅäÖÃ
    lstr = Trim(Left(s, i - 1))
    rstr = Trim(Right(s, Len(s) - i))
    If lstr = "" Or rstr = "" Then
        Exit Sub
    End If
    Select Case UCase(lstr)
        Case "#TITLE"
            rstr = rstr
        Case "#QUOTATION1"
            m_sQuotFlag1 = rstr
        Case "#QUOTATION2"
            m_sQuotFlag2 = rstr
        Case "#CONTINUE_QUOTE"
            m_bMultiLine = IIf(LCase(rstr) = "y", True, False)
        Case "#LINECOMMENT"
            m_sLnCommFlag1 = rstr
        Case "#LINECOMMENT2"
            m_sLnCommFlag2 = rstr
        Case "#COMMENTON"
            m_sBlkCommHead1 = rstr
        Case "#COMMENTOFF"
            m_sBlkCommFoot1 = rstr
        Case "#COMMENTON2"
            m_sBlkCommHead2 = rstr
        Case "#COMMENTOFF2"
            m_sBlkCommFoot2 = rstr
        Case "#CASE"
            m_bCase = IIf(LCase(rstr) = "y", True, False)
            
        Case "#ESCAPE"
            m_sEscape = rstr
        Case Else
           
    End Select
End Sub

Private Sub ReadFont(sValue As String, i As Integer, Font As CFont)
'*¶ÁÈ¡ÎÄ¼þÖÐµÄ×ÖÌåÉèÖÃ
'*      i       --"="ËùÔÚÎ»ÖÃ
'*      font    --Òª´«ÈëµÄ×ÖÌåÀà

Dim lstr        As String
Dim rstr        As String
    lstr = Trim(Left(sValue, i - 1))
    rstr = Trim(Right(sValue, Len(sValue) - i))
    If lstr <> "" And rstr <> "" Then
        Select Case LCase(lstr)
            Case "forecolor"
                Font.bForecolor = True
                Font.ForeColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "backcolor"
                Font.bBackcolor = True
                Font.BackColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "bold"
                Font.bold = IIf(LCase(rstr) = "y", True, False)
            Case "italic"
                Font.italic = IIf(LCase(rstr) = "y", True, False)
            Case "underline"
                Font.underline = IIf(LCase(rstr) = "y", True, False)
            Case "strike"
                Font.strike = IIf(LCase(rstr) = "y", True, False)
            Case Else
                
        End Select
    End If
End Sub

Private Sub Class_Initialize()
    Set m_objColorManager = New CColor
    Set m_objCommFont = New CFont
    Set m_objQuotFont = New CFont
    Set m_objNumFont = New CFont
End Sub

Private Sub Class_Terminate()
Dim i   As Integer
    Set m_objColorManager = Nothing
    Set m_objCommFont = Nothing
    Set m_objQuotFont = Nothing
    Set m_objNumFont = Nothing
    
    For i = 1 To m_iKwdCount
        Set m_objKwdFont(i) = Nothing
    Next i
    
    Erase m_objKwdFont
    Erase m_sKwdTitle
    Erase m_sKwdSet
    Erase m_sKwdRtfHead
    Erase m_sKwdRtfFoot
    Erase m_sWord
End Sub

'**************************************************************
'*HighLightRichEdit
'*¹¦ÄÜ£º¶Ô¿Ø¼þ½øÐÐÓï·¨¸ßÁÁÏÔÊ¾
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    ctrl      ÀàÐÍ£ºRichTextBox
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-09-29  17:38:22
'**************************************************************
Public Sub HighLightRichEdit(ctrl As RichTextBox)
Dim sFileName   As String

    sFileName = TempFileName
    HighLightCtrl ctrl, sFileName, True
    
End Sub


'**************************************************************
'*HighLightCtrl
'*¹¦ÄÜ£º¶Ô¿Ø¼þ½øÐÐÓï·¨¸ßÁÁÏÔÊ¾
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    ctrl           ÀàÐÍ£ºRichTextBox      --¿Ø¼þ
'*    filename       ÀàÐÍ£ºString           --ÁÙÊ±ÎÄ¼þÃû
'*    bKillFile      ÀàÐÍ£ºBoolean          --Íê³ÉºóÊÇ·ñÉ¾³ýÎÄ¼þ
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-09-29  17:38:22
'**************************************************************
Private Sub HighLightCtrl(ctrl As RichTextBox, filename As String, bKillFile As Boolean)
Dim sLine() As String
Dim i       As Long
Dim l       As Long
Dim sTmp        As String
Dim iProgress   As Integer      '*µ±Ç°´¦Àí½ø¶È
Dim iTmp        As Integer

    
    sTmp = ctrl.Text
    l = Len(sTmp)
    If l = 0 Then
        Exit Sub
    End If
    '*ÏÈ½«
    Open filename For Output As #1
    
        Print #1, HEAD_FIRST & m_sColorRtf & HEAD_LAST
 
        m_iQuotState = 0
        m_iCommState = 0
        
        iProgress = 0
        
        '*¶ÔÓÚ´óÎÄ±¾½øÐÐ·Ö¿é´¦Àí£¬Ã¿1KÎªÒ»¸öµ¥Î»
        Dim aTmp()      As String
        Dim j           As Long
        Dim lstPos      As Long
        lstPos = 1
        ReDim aTmp(1 To l / 4000 + 1)
        For i = 1 To UBound(aTmp)
            j = InStr(lstPos + 4000, sTmp, vbCrLf, vbTextCompare)
            If j = 0 Then
                j = l
                aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
                lstPos = -1
            Else
                aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
                lstPos = j + 1
            End If
            '*´¦Àí¹¦ÄÜ×Ö·ûºÍ·Ö¸ô×Ö·û
            aTmp(i) = Replace(aTmp(i), "\", "\\")
            aTmp(i) = Replace(aTmp(i), "{", "\{")
            aTmp(i) = Replace(aTmp(i), "}", "\}")
            aTmp(i) = Replace(aTmp(i), "(", " ( ")
            aTmp(i) = Replace(aTmp(i), ")", " ) ")
            aTmp(i) = Replace(aTmp(i), ";", " ; ")
            aTmp(i) = Replace(aTmp(i), "#", " # ")
            aTmp(i) = Replace(aTmp(i), Chr(9), " " & Chr(9) & " ")
            aTmp(i) = Replace(aTmp(i), vbTab, " " & vbTab & " ")
        
            sLine = Split(aTmp(i), vbCrLf)
            
                
            For j = 0 To UBound(sLine)
            
                Call HighLightLine(sLine(j))    '*´¦ÀíÒ»ÐÐ
                
                '*»Ö¸´·Ö¸ô×Ö·û
                aTmp(i) = Replace(sLine(j), " ( ", "(")
                aTmp(i) = Replace(aTmp(i), " ) ", ")")
                aTmp(i) = Replace(aTmp(i), " ; ", ";")
                aTmp(i) = Replace(aTmp(i), " # ", "#")
                aTmp(i) = Replace(aTmp(i), " " & Chr(9) & " ", Chr(9))
                aTmp(i) = Replace(aTmp(i), " " & vbTab & " ", vbTab)
                
                Print #1, aTmp(i) & "\par "
                
                
    
            Next j
            
            If lstPos = -1 Then
                Exit For
            End If
            
            iTmp = i * 100 / UBound(aTmp)
            If iProgress <> iTmp Then
                iProgress = iTmp
                RaiseEvent Progress(iProgress)
            End If
                
        Next i
        
        Print #1, "}"

    Close #1

    ctrl.LoadFile filename, rtfRTF
    
    Erase sLine
    
    If Dir(filename) <> "" And bKillFile Then
        Kill filename
    End If
End Sub

Private Sub HighLightLine(ByRef sLine As String)
'*É¨Ãè×Ö·û£¬Êä³ö¸ñÊ½»¯ºóµÄ×Ö·û´®
Dim i       As Long
Dim j       As Integer
Dim k       As Integer
Dim bComm As Boolean

    m_sWord = Split(sLine, " ")
    sLine = ""
    
    k = UBound(m_sWord)
    
    For i = 0 To k

        '*Èç¹û²»ÊÇ×Ö·û´®£¬Ôò½øÐÐÏÈ½øÐÐ×¢ÊÍÅÐ¶Ï£¬ÔÙ½øÐÐ¹Ø¼ü×ÖÅÐ¶Ï
        If m_iQuotState = 0 Then
        
            bComm = TreatComment(m_sWord(i), (i = k))         '*´¦Àí×¢ÊÍ
        
        End If

        If Not bComm Then

                If TreatQuot(m_sWord(i), (i = k), 1) Then       '*´¦Àí×Ö·û´®
                
                    '*»Ö¸´±»´¦Àí¹ýµÄ×Ö·û´®±êÖ¾
                    If m_sQuotFlag1 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(1), m_sQuotFlag1)
                    End If
                    If m_sQuotFlag2 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(2), m_sQuotFlag2)
                    End If
                Else
                
                    If m_iQuotState = 0 Then
                        Call TreatKeyWord(m_sWord(i))         '*´¦Àí¹Ø¼ü×Ö
                    End If
                    
                End If

        End If
        
        If i = k Then
            sLine = sLine & m_sWord(i)
        Else
            sLine = sLine & m_sWord(i) & " "
        End If
        
    Next i

End Sub

Private Function TreatComment(ByRef sWord As String, bLast As Boolean) As Boolean
'*´¦Àí×¢ÊÍ
'*      bLast       --ÊÇ·ñ×îºóÒ»¸ö×Ö·û
'*      ·µ»ØÖµ£º    sWordÊÇ·ñ×¢ÊÍ
    Select Case m_iCommState
        Case 0
            If m_sLnCommFlag1 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag1, vbBinaryCompare) = 1 Then  '*ÊÇµ¥ÐÐ×¢ÊÍÒ»
                    m_iCommState = 1
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Í·
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*Èç¹û×îºóÒ»ÐÐ£¬ÔÙ½øÐÐ´¦Àí
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sLnCommFlag2 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag2, vbBinaryCompare) = 1 Then  '*ÊÇµ¥ÐÐ×¢ÊÍ¶þ
                    m_iCommState = 2
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Í·
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*Èç¹û×îºóÒ»¸ö´Ê£¬ÔÙ½øÐÐ´¦Àí
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead1, vbBinaryCompare) = 1 Then  '*ÊÇ¶àÐÐ×¢ÊÍÒ»
                    m_iCommState = 3
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Í·
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead2, vbBinaryCompare) = 1 Then '*ÊÇ¶àÐÐ×¢ÊÍ¶þ
                    m_iCommState = 4
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Í·
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
        Case 1, 2
            TreatComment = True
            If bLast Then       '*µ¥ÐÐ×¢ÊÍ£¬ÇÒµ±Ç°Îª´ËÐÐ×îºóÒ»¸ö´Ê£¬½«×¢ÊÍ±êÖ¾»Ö¸´Îª0
                sWord = sWord & m_sCommRtfFoot
                m_iCommState = 0
            End If
        Case 3
            TreatComment = True
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot1, , vbBinaryCompare) <> 0 Then '*ÊÇ¶àÐÐ×¢ÊÍÒ»
                    m_iCommState = 0
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Î²
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case 4
            TreatComment = True
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot2, , vbBinaryCompare) <> 0 Then '*ÊÇ¶àÐÐ×¢ÊÍ¶þ
                    m_iCommState = 0
                    '*¼ÓÉÏ×¢ÊÍ×ÖÌåRTF¸ñÊ½Î²
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case Else
        
    End Select
    
End Function

Private Function TreatKeyWord(ByRef sWord As String) As Boolean
'*´¦Àí¹Ø¼ü×Ö
'*      ·µ»ØÖµ£º    sWordÊÇ·ñ¹Ø¼ü×Ö
Dim i           As Integer
Dim bKeyWord    As Boolean
    '* ÏÈ´¦ÀíÊÇ²»ÊÇÊý×Ö
    If IsNumeric(sWord) Then
        sWord = m_sNumRtfHead & sWord & m_sNumRtfFoot
        Exit Function
    End If
    
    For i = 1 To m_iKwdCount
        If m_bCase Then
            If (InStr(1, m_sKwdSet(i), Chr(3) & sWord & Chr(3), vbBinaryCompare) <> 0) Then
               bKeyWord = True
            End If
        Else
            If (InStr(1, m_sKwdSet(i), Chr(3) & UCase(sWord) & Chr(3), vbBinaryCompare) <> 0) Then
                bKeyWord = True
            End If
        End If
        
        If bKeyWord Then
            '*¼ÓÉÏ×ÖÌåRTF¸ñÊ½µÄÍ·Î²ÓÚ´ÊÁ½¶Ë
            sWord = m_sKwdRtfHead(i) & sWord & m_sKwdRtfFoot(i)
            TreatKeyWord = True
            Exit Function
        End If
    Next i
End Function

Private Function TreatQuot(ByRef sWord As String, bLast As Boolean, iBegin As Integer) As Boolean
'*´¦Àí×Ö·û´®
'*      bLast       ÊÇ·ñ´ËÐÐ×îºóÒ»¸ö´Ê
'*      iBegin      ËÑË÷Î»ÖÃ
'*      ·µ»ØÖµ£º    ÊÇ·ñ×ö¹ý×Ö·û´®´¦Àí
Dim i       As Integer
Dim j       As Integer
Dim sQuot   As String       '*µ±Ç°ÓÃÀ´ÅÐ¶ÏµÄ×Ö·û´®±ß½ç·û
    Select Case m_iQuotState
        Case 0
            '*ÅÐ¶ÏÄÄÖÖ×Ö·û´®±êÖ¾ÏÈ³öÏÖ£¬Ê¹ÓÃÏÈ³öÏÖµÄ×Ö·û´®±êÖ¾À´¾ö¶¨×Ö·û´®ÆðÊ¼Î»ÖÃ
            i = 0
            j = 0
            If m_sQuotFlag1 <> "" Then
                i = InStr(iBegin, sWord, m_sQuotFlag1, vbBinaryCompare)
            End If
            If m_sQuotFlag2 <> "" Then
                j = InStr(iBegin, sWord, m_sQuotFlag2, vbBinaryCompare)
            End If
            If i = 0 And j = 0 Then     '*Ã»ÓÐ×Ö·û´®±êÖ¾³öÏÖ£¬²»ÓÃÅÐ¶Ï
                Exit Function
            End If
            If (j = 0) Or (i <> 0 And i < j) Then
                sQuot = m_sQuotFlag1
            End If
            If (j <> 0 And j < i) Or (i = 0) Then
                sQuot = m_sQuotFlag2
            End If
        Case 1
            sQuot = m_sQuotFlag1
        Case 2
            sQuot = m_sQuotFlag2
        Case Else
    End Select
    
    '*²éÕÒµ±Ç°µÄ×Ö·û´®±êÖ¾
    i = InStr(iBegin, sWord, sQuot, vbBinaryCompare)
    
    '*Èç¹ûÇ°Ò»¸ö×Ö·ûÊÇÃô¸Ð×Ö·ûÇ°×º£¬Ôò´Ë´Î²éÕÒÎÞÐ§
    If m_sEscape <> "" Then
        If i > 1 Then
            j = InStr(iBegin, sWord, m_sEscape)
            If j <> 0 And j = i - Len(m_sEscape) Then
                '*½«´Ë×Ö·û´®±êÖ¾ÆÁ±Ü
                If sQuot = m_sQuotFlag1 Then
                    sWord = Replace(sWord, sQuot, Chr(1), , 1)
                Else
                    sWord = Replace(sWord, sQuot, Chr(2), , 1)
                End If
                TreatQuot = True
                '*¼ÌÐøµÝ¹é´¦Àí
                Call TreatQuot(sWord, bLast, i + 1)
                Exit Function
            End If
        End If
    End If
    
    If i <> 0 Then
    
        TreatQuot = True
        
        '*¸ñÊ½»¯×Ö·û´®
        If sQuot = m_sQuotFlag1 Then
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(1), , 1)
                m_iQuotState = 1
            Else
                sWord = Replace(sWord, sQuot, Chr(1) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        Else
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(2), , 1)
                m_iQuotState = 2
            Else
                sWord = Replace(sWord, sQuot, Chr(2) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        End If

        '*¼ÌÐøµÝ¹é´¦Àí
        Call TreatQuot(sWord, bLast, i)
        
    Else
    
        '*Èç¹ûÒÑ¾­Ã»ÓÐÁË±êÖ¾£¬µ«ÐøÐÐÎÞÐ§µÄÇé¿öÏÂ£¬µ±Ç°Õâ¸ö´ÊÎª×îºóÒ»¸öÊ±£¬Ç¿ÐÐ½áÊø
        If bLast And Not m_bMultiLine And m_iQuotState <> 0 Then
            sWord = sWord & m_sQuotRtfFoot
            m_iQuotState = 0
        End If
        
    End If
    

End Function


'**************************************************************
'*Rtf2Html
'*¹¦ÄÜ£º    ½«RTF¿Ø¼þÄÚÈÝ×ª»»ÎªHtml×Ö·û´®
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    rTxt      ÀàÐÍ£ºRichTextBox
'*·µ»Ø²ÎÊý£º
'*      HTML×Ö·û´®
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Html(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False
    
    '*¶ÁÈëÎÄ¼þ
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Html = StrConv(byt, vbUnicode)
    
    '*½«ÌØÊâ×Ö·û»Ö¸´
    Rtf2Html = Replace(Rtf2Html, "\\", "\")
    Rtf2Html = Replace(Rtf2Html, "\{", "{")
    Rtf2Html = Replace(Rtf2Html, "\}", "}")
    
    '*ÏÈÉ¾³ýRTFÎÄ¼þÍ·
    Rtf2Html = Replace(Rtf2Html, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*×ª»»»»ÐÐ·ûºÅ
    Rtf2Html = Replace(Rtf2Html, "\par ", "")
    
    '*½«ÌØÊâ×Ö·û´¦Àí
    Rtf2Html = Replace(Rtf2Html, "&", "&amp;")
    Rtf2Html = Replace(Rtf2Html, "<", "&lt;")
    Rtf2Html = Replace(Rtf2Html, ">", "&gt;")
    '*×ª»»×¢ÊÍ×ÖÌå·ûºÅ
    FontRtf2Html Rtf2Html, m_objQuotFont
    FontRtf2Html Rtf2Html, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Html Rtf2Html, m_objKwdFont(i)
    Next i
    
    '*È¥³ý×îºóÒ»¸ö"}"
    Rtf2Html = Left(Rtf2Html, Len(Rtf2Html) - 4)
    
    
    '*¼ÓÉÏHTMLÍ·Î²
    Rtf2Html = HEAD_HTML & Rtf2Html & TAIL_HTML
 
End Function

Private Sub FontRtf2Html(ByRef sDeal As String, ByRef objFont As CFont)
'*¶ÔsDeal×Ö·û´®½øÐÐRTFµ½HTMLµÄ×ÖÌå×ª»»
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getHtmlHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getHtmlTail)
End Sub

'**************************************************************
'*Rtf2Ubb
'*¹¦ÄÜ£º    ½«RTF¿Ø¼þÄÚÈÝ×ª»»ÎªUbb×Ö·û´®
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    rTxt      ÀàÐÍ£ºRichTextBox
'*·µ»Ø²ÎÊý£º
'*      HTML×Ö·û´®
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Ubb(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False

    '*¶ÁÈëÎÄ¼þ
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Ubb = StrConv(byt, vbUnicode)
    
    '*½«ÌØÊâ×Ö·û»Ö¸´
    Rtf2Ubb = Replace(Rtf2Ubb, "\\", "\")
    Rtf2Ubb = Replace(Rtf2Ubb, "\{", "{")
    Rtf2Ubb = Replace(Rtf2Ubb, "\}", "}")
    
    '*ÏÈÉ¾³ýRTFÎÄ¼þÍ·
    Rtf2Ubb = Replace(Rtf2Ubb, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*×ª»»»»ÐÐ·ûºÅ
    Rtf2Ubb = Replace(Rtf2Ubb, "\par ", vbCrLf)
    
    '*×ª»»×¢ÊÍ×ÖÌå·ûºÅ
    FontRtf2Ubb Rtf2Ubb, m_objQuotFont
    FontRtf2Ubb Rtf2Ubb, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Ubb Rtf2Ubb, m_objKwdFont(i)
    Next i
    
    '*È¥³ý×îºóÒ»¸ö"}"
    Rtf2Ubb = Left(Rtf2Ubb, Len(Rtf2Ubb) - 4)
    
 
End Function

Private Sub FontRtf2Ubb(ByRef sDeal As String, ByRef objFont As CFont)
'*¶ÔsDeal×Ö·û´®½øÐÐRTFµ½UBBµÄ×ÖÌå×ª»»
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getUBBHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getUBBTail)
End Sub
