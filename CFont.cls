VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'*ÀàÄ£¿éÃû³Æ£ºCFont
'*ÀàÄ£¿é¹¦ÄÜ£ºµÃµ½RTF¸ñÊ½ÄÚ×ÖÌåÐÅÏ¢
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-09-29  13:32:49
'***************************************************************
Option Explicit

Public ForeColor    As OLE_COLOR
Public BackColor    As OLE_COLOR

Public bForecolor   As Boolean
Public bBackcolor   As Boolean

Public bold         As Boolean
Public italic       As Boolean
Public underline    As Boolean
Public strike       As Boolean

'**************************************************************
'*getRtfHead
'*¹¦ÄÜ£ºµÃµ½RTF¸ñÊ½ÖÐ×ÖÌåÇ°²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-09-29  13:36:18
'**************************************************************
Public Function getRtfHead(ByRef color As CColor) As String
    If bBackcolor Then
        getRtfHead = getRtfHead & "\highlight" & color.getColorPos(BackColor)
    End If
    If bForecolor Then
        getRtfHead = getRtfHead & "\cf" & color.getColorPos(ForeColor)
    End If
    If bold Then
        getRtfHead = getRtfHead & "\b"
    End If
    If italic Then
        getRtfHead = getRtfHead & "\i"
    End If
    If underline Then
        getRtfHead = getRtfHead & "\ul"
    End If
    If strike Then
        getRtfHead = getRtfHead & "\strike"
    End If
    
    If getRtfHead <> "" Then
        getRtfHead = getRtfHead & " "
    End If
End Function

'**************************************************************
'*getRtfTail
'*¹¦ÄÜ£ºµÃµ½RTF¸ñÊ½ÖÐ×ÖÌåºó²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-09-29  13:36:18
'**************************************************************
Public Function getRtfTail() As String
    If strike Then
        getRtfTail = getRtfTail & "\strike0"
    End If
    If underline Then
        getRtfTail = getRtfTail & "\ulnone"
    End If
    If italic Then
        getRtfTail = getRtfTail & "\i0"
    End If
    If bold Then
        getRtfTail = getRtfTail & "\b0"
    End If
    If bForecolor Then
        getRtfTail = getRtfTail & "\cf0"
    End If
    If bBackcolor Then
        getRtfTail = getRtfTail & "\highlight0"
    End If

    If getRtfTail <> "" Then
        getRtfTail = getRtfTail & " "
    End If
End Function


'**************************************************************
'*getHtmlHead
'*¹¦ÄÜ£º    µÃµ½HTML¸ñÊ½ÖÐ×ÖÌåÇ°²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    color      ÀàÐÍ£ºCColor
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:59:01
'**************************************************************
Public Function getHtmlHead(ByRef color As CColor) As String
    If strike Then
        getHtmlHead = getHtmlHead & "<STRIKE>"
    End If
    If underline Then
        getHtmlHead = getHtmlHead & "<U>"
    End If
    If italic Then
        getHtmlHead = getHtmlHead & "<I>"
    End If
    If bold Then
        getHtmlHead = getHtmlHead & "<B>"
    End If
    If bForecolor Then
        getHtmlHead = getHtmlHead & "<FONT COLOR = ""#" & color.getColorHex(ForeColor) & """>"
    End If
    If bBackcolor Then
        getHtmlHead = getHtmlHead & "<span style=""background-color: #" _
                                  & color.getColorHex(BackColor) & """>"
    End If

End Function


'**************************************************************
'*getHtmlTail
'*¹¦ÄÜ£º    µÃµ½HTML¸ñÊ½ÖÐ×ÖÌåºó²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:50:36
'**************************************************************
Public Function getHtmlTail() As String
    If strike Then
        getHtmlTail = getHtmlTail & "</STRIKE>"
    End If
    If underline Then
        getHtmlTail = getHtmlTail & "</U>"
    End If
    If italic Then
        getHtmlTail = getHtmlTail & "</I>"
    End If
    If bold Then
        getHtmlTail = getHtmlTail & "</B>"
    End If
    If bForecolor Then
        getHtmlTail = getHtmlTail & "</FONT>"
    End If
    If bBackcolor Then
        getHtmlTail = getHtmlTail & "</span>"
    End If

End Function


'**************************************************************
'*getUBBHead
'*¹¦ÄÜ£º    µÃµ½UBB¸ñÊ½ÖÐ×ÖÌåÇ°²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*    color      ÀàÐÍ£ºCColor
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:59:01
'**************************************************************
Public Function getUBBHead(ByRef color As CColor) As String
    If bForecolor Then
        getUBBHead = getUBBHead & "[color=#" & color.getColorHex(ForeColor) & "]"
    End If
    If underline Then
        getUBBHead = getUBBHead & "[U]"
    End If
    If italic Then
        getUBBHead = getUBBHead & "[I]"
    End If
    If bold Then
        getUBBHead = getUBBHead & "[B]"
    End If

End Function


'**************************************************************
'*getUBBTail
'*¹¦ÄÜ£º    µÃµ½UBB¸ñÊ½ÖÐ×ÖÌåºó²¿
'*ËµÃ÷£º
'*´«Èë²ÎÊý£º
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame  ÈÕÆÚ£º2002-10-17  17:50:36
'**************************************************************
Public Function getUBBTail() As String
    If bold Then
        getUBBTail = getUBBTail & "[/B]"
    End If
    If italic Then
        getUBBTail = getUBBTail & "[/I]"
    End If
    If underline Then
        getUBBTail = getUBBTail & "[/U]"
    End If
    If bForecolor Then
        getUBBTail = getUBBTail & "[/color]"
    End If
End Function


Private Sub Class_Initialize()
    ForeColor = vbBlack
    BackColor = vbWhite
End Sub
