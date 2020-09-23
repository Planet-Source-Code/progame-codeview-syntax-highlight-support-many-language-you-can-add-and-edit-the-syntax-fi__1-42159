VERSION 5.00
Begin VB.UserControl FlatProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ScaleHeight     =   345
   ScaleWidth      =   3855
   Begin VB.Label lab1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1755
      TabIndex        =   0
      Top             =   90
      Width           =   555
   End
   Begin VB.Label lab2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1755
      TabIndex        =   1
      Top             =   45
      Width           =   540
   End
End
Attribute VB_Name = "FlatProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'*¿Ø¼þÃû³Æ£ºFlatProgressBar
'*¿Ø¼þ¹¦ÄÜ£ºÆ½»¬½ø¶ÈÌõ
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-10-07  17:57:22
'***************************************************************
Option Explicit
'È±Ê¡ÊôÐÔÖµ:
Const m_def_Text = ""
Const m_def_ShowNumber = True
Const m_def_NumberColor = &HFFFFFF
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
'ÊôÐÔ±äÁ¿:
Dim m_Text As String
Dim m_ShowNumber As Boolean
Dim m_NumberColor As OLE_COLOR
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_Value As Integer



'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "·µ»Ø/ÉèÖÃ¶ÔÏóÖÐÎÄ±¾ºÍÍ¼ÐÎµÄ±³¾°É«¡£"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "·µ»Ø/ÉèÖÃ¶ÔÏóÖÐÎÄ±¾ºÍÍ¼ÐÎµÄÇ°¾°É«¡£"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Call Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Ç¿ÖÆÍêÈ«ÖØ»­Ò»¸ö¶ÔÏó¡£"
    UserControl.Refresh
End Sub

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
'*Min±ØÐë´óÓÚµÈÓÚ0£¬Ð¡ÓÚMax
    If New_Min < 0 Or New_Min >= m_Max Then
        Err.Raise vbObjectError + 513, "", "Invalid Value"
        Exit Property
    End If
    m_Min = New_Min
    PropertyChanged "Min"
    Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
'*Max±ØÐë´óÓÚMin
    If New_Max <= m_Min Then
        Err.Raise vbObjectError + 513, "", "Invalid Value"
        Exit Property
    End If
    m_Max = New_Max
    PropertyChanged "Max"
    Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=7,0,2,0
Public Property Get Value() As Integer
Attribute Value.VB_MemberFlags = "400"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
'*Value±ØÐëÔÚMinµ½MaxÖ®¼ä
    If New_Value < m_Min Or New_Value > m_Max Then
        Err.Raise vbObjectError + 513, "", "Invalid Value"
        Exit Property
    End If
    If Ambient.UserMode = False Then Err.Raise 387
    m_Value = New_Value
    PropertyChanged "Value"
    Draw
End Property

'ÎªÓÃ»§¿Ø¼þ³õÊ¼»¯ÊôÐÔ
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    Set UserControl.Font = Ambient.Font
    m_ShowNumber = m_def_ShowNumber
    m_NumberColor = m_def_NumberColor
    m_Text = m_def_Text
End Sub

'´Ó´æÖüÆ÷ÖÐ¼ÓÔØÊôÐÔÖµ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ShowNumber = PropBag.ReadProperty("ShowNumber", m_def_ShowNumber)
    m_NumberColor = PropBag.ReadProperty("NumberColor", m_def_NumberColor)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
End Sub

Private Sub UserControl_Resize()
    Draw
End Sub

'½«ÊôÐÔÖµÐ´µ½´æ´¢Æ÷
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ShowNumber", m_ShowNumber, m_def_ShowNumber)
    Call PropBag.WriteProperty("NumberColor", m_NumberColor, m_def_NumberColor)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
End Sub

Private Sub Draw()
'*»æÖÆ¿Ø¼þ

    DrawNumber
    DrawLine

    
End Sub

Private Sub DrawLine()
'*»æÖÆ±ßÏßºÍ½ø¶ÈÌõ
Dim lWidth      As Long     '*½ø¶ÈÌõ¿í¶È
    lWidth = (UserControl.Width - 40) * m_Value / 100
    UserControl.Cls
    UserControl.Line (20, 20)-Step(lWidth, UserControl.Height - 40), , BF
    UserControl.Line (0, 0)-(UserControl.Width, 0), vb3DShadow
    UserControl.Line (0, 0)-(0, UserControl.Height), vb3DShadow
    UserControl.Line (0, UserControl.Height - 20)-Step(UserControl.Width, 0), vbWhite
    UserControl.Line (UserControl.Width - 20, 0)-Step(0, UserControl.Height), vbWhite
End Sub

Private Sub DrawNumber()
'*»æÖÆÊýÖµ
Dim lWidth      As Long     '*½ø¶ÈÌõ¿í¶È
    lWidth = (UserControl.Width - 40) * m_Value / 100
    '*ÏÔÊ¾ÊýÖµ´¦Àí
    If Not ShowNumber Then
        lab1.Visible = False
        lab2.Visible = False
        Exit Sub
    End If
    '*½«lab±ä³ÉÏàÍ¬×ÖÌå
    With lab1.Font
        .bold = UserControl.Font.bold
        .Charset = UserControl.Font.Charset
        .italic = UserControl.Font.italic
        .Name = UserControl.Font.Name
        .Size = UserControl.Font.Size
        .Strikethrough = UserControl.Font.Strikethrough
        .underline = UserControl.Font.underline
        .Weight = UserControl.Font.Weight
    End With
    With lab2.Font
        .bold = UserControl.Font.bold
        .Charset = UserControl.Font.Charset
        .italic = UserControl.Font.italic
        .Name = UserControl.Font.Name
        .Size = UserControl.Font.Size
        .Strikethrough = UserControl.Font.Strikethrough
        .underline = UserControl.Font.underline
        .Weight = UserControl.Font.Weight
    End With
    lab1.AutoSize = True
    lab2.AutoSize = True
    If m_Text = "" Then
        lab1.Caption = m_Value & "%"
    Else
        lab1.Caption = m_Text
    End If
    lab2.Caption = lab1.Caption
    lab1.BackColor = UserControl.ForeColor
    lab2.BackColor = UserControl.BackColor
    lab2.Visible = True
    lab1.Left = (UserControl.Width - lab2.Width) / 2
    lab1.Top = (UserControl.Height - lab1.Height) / 2
    lab2.Top = lab1.Top
    lab2.Left = lab1.Left
    lab1.Height = lab2.Height
    lab1.AutoSize = False
    If (lWidth + 20) > (UserControl.Width - lab2.Width) / 2 Then
        lab1.Width = (lWidth + 20) - (UserControl.Width - lab2.Width) / 2
        If lab1.Width > lab2.Width Then
            lab1.Width = lab2.Width
            'lab2.Visible = False
        Else
            'lab2.Visible = True
        End If
        lab1.Visible = True
    Else
        lab1.Visible = False
    End If

    lab1.ForeColor = &HFFFFFF Xor m_NumberColor
    lab2.ForeColor = m_NumberColor
    
    lab1.ZOrder 0
End Sub

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "·µ»ØÒ»¸ö Font ¶ÔÏó¡£"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=0,0,0,0
Public Property Get ShowNumber() As Boolean
    ShowNumber = m_ShowNumber
End Property

Public Property Let ShowNumber(ByVal New_ShowNumber As Boolean)
    m_ShowNumber = New_ShowNumber
    PropertyChanged "ShowNumber"
    Draw
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=10,0,0,&Hffffff
Public Property Get NumberColor() As OLE_COLOR
    NumberColor = m_NumberColor
End Property

Public Property Let NumberColor(ByVal New_NumberColor As OLE_COLOR)
    m_NumberColor = New_NumberColor
    PropertyChanged "NumberColor"
End Property

'×¢Òâ£¡²»ÒªÉ¾³ý»òÐÞ¸ÄÏÂÁÐ±»×¢ÊÍµÄÐÐ£¡
'MemberInfo=13,0,0,
Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    Draw
End Property

