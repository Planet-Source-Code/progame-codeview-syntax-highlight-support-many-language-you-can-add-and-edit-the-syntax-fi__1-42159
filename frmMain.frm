VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "CodeView"
   ClientHeight    =   5565
   ClientLeft      =   1470
   ClientTop       =   2115
   ClientWidth     =   9330
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   9330
   Begin CodeView.FlatProgressBar prg 
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   5310
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   423
      ForeColor       =   7021576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberColor     =   0
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   990
      Top             =   3915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar staInfo 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   12417
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTXT 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTXTBack 
      Height          =   240
      Left            =   8190
      TabIndex        =   1
      Top             =   3285
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":03A0
   End
   Begin VB.TextBox txtInfo 
      Height          =   465
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   6120
      Width           =   465
   End
   Begin VB.Label labInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   495
      TabIndex        =   2
      Top             =   5220
      Width           =   90
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_file_open 
         Caption         =   "Open..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_file_save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_file_saveas 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu mnu_file_ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_file_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_syx 
      Caption         =   "&Syntax"
      Begin VB.Menu mnu_syx_item 
         Caption         =   "&Normal File"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnu_syx_item 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnu_syx_item 
         Caption         =   "Visual &Basic"
         Index           =   2
      End
      Begin VB.Menu mnu_syx_item 
         Caption         =   "Visual &C++"
         Index           =   3
      End
      Begin VB.Menu mnu_syx_item 
         Caption         =   "Object &Pascal"
         Index           =   4
      End
      Begin VB.Menu mnu_syx_item 
         Caption         =   "Transact &SQL"
         Index           =   5
      End
   End
   Begin VB.Menu mnu_tool 
      Caption         =   "&Tool"
      Begin VB.Menu mnu_tool_html 
         Caption         =   "Convert to &Html"
      End
      Begin VB.Menu mnu_tool_ubb 
         Caption         =   "Convert to &UBB"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_help_about 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*´°ÌåÃû³Æ£ºfrmMain
'*´°Ìå¹¦ÄÜ£º´úÂë²é¿´
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-10-07  17:56:01
'***************************************************************
Option Explicit


Private WithEvents Syntax   As CSyntax
Attribute Syntax.VB_VarHelpID = -1

Private Sub Form_Load()
    staInfo.Panels(2).Text = Replace(mnu_syx_item(0).Caption, "&", "")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    RTXT.Width = Me.Width - 120
    RTXT.Height = Me.Height - staInfo.Height - 700
    prg.Top = Me.Height - staInfo.Height - 650
    prg.ZOrder 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RTXT.Text = ""
    Set Syntax = Nothing
    End
End Sub


Private Sub mnu_file_exit_Click()
    Unload Me
End Sub

Private Sub mnu_file_open_Click()
'*Open File
    With dlgFile
        .CancelError = True
        On Error GoTo Err_Proc
        .DialogTitle = "Open File"
        .ShowOpen
        RTXT.LoadFile .filename, rtfText
        Call Highlight

    End With

    Exit Sub
    
Err_Proc:
End Sub

Private Sub mnu_file_save_Click()
    Screen.MousePointer = vbHourglass
    RTXT.SaveFile RTXT.filename, rtfText
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_help_about_Click()
    MsgBox "CodeView" & vbCrLf & vbCrLf & "      --by progame  "
End Sub

Private Sub mnu_syx_item_Click(Index As Integer)
'*¸Ä±äÓï·¨ÎÄ¼þ
Dim i       As Integer
    For i = 0 To 5
        mnu_syx_item(i).Checked = False
    Next i
    mnu_syx_item(Index).Checked = True
    
    Set Syntax = Nothing
    Set Syntax = New CSyntax
    
    Select Case Index
        Case 2
            Syntax.ReadFile App.Path & "\stx\vb.stx"
        Case 3
            Syntax.ReadFile App.Path & "\stx\cpp.stx"
        Case 4
            Syntax.ReadFile App.Path & "\stx\pas.stx"
        Case 5
            Syntax.ReadFile App.Path & "\stx\tsql.stx"
    End Select
    
    If Index > 1 Then
        Highlight
    End If
    
    staInfo.Panels(2).Text = Replace(mnu_syx_item(Index).Caption, "&", "")
    
End Sub


Private Sub Highlight()
Dim t       As Single
Dim lPos    As Long

    If Syntax Is Nothing Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
    lPos = RTXT.SelStart

    prg.Visible = True
    t = Timer
    Syntax.HighLightRichEdit RTXT
    'MsgBox Timer - t
    prg.Visible = False
    
    On Error Resume Next
    
    RTXT.SelStart = lPos
    staInfo.Panels(1).Text = "File Length:" & Len(RTXT.Text) & "  Use Seconds:" & Timer - t
    RTXT.SetFocus
    
    'txtInfo.Text = RTXT.TextRTF
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnu_tool_Click()
'*Èç¹ûµ±Ç°Ã»ÓÐÓï·¨ÉèÖÃ£¬ÔòÎÞÐ§
    mnu_tool_html.Enabled = True
    mnu_tool_ubb.Enabled = True
    If mnu_syx_item(0).Checked Then
        mnu_tool_html.Enabled = False
        mnu_tool_ubb.Enabled = False
    End If
End Sub

Private Sub mnu_tool_html_Click()
'*×ª»»HTML×Ö·û´®
Dim fHtml       As frmHtml

    Set fHtml = New frmHtml
        With fHtml
            .txtHtml.SelRTF = RTXT.SelRTF
            .txtHtml.Text = Syntax.Rtf2Html(.txtHtml)
            .Caption = "Convert to HTML"
            .Show vbModal, Me
        End With
    Set fHtml = Nothing
End Sub

Private Sub mnu_tool_ubb_Click()
'*×ª»»UBB×Ö·û´®
Dim fHtml       As frmHtml

    Set fHtml = New frmHtml
        With fHtml
            .txtHtml.SelRTF = RTXT.SelRTF
            .txtHtml.Text = Syntax.Rtf2Ubb(.txtHtml)
            .Caption = "Convert to UBB"
            .Show vbModal, Me
        End With
    Set fHtml = Nothing
End Sub

Private Sub Syntax_Progress(Value As Integer)
    prg.Value = Value
End Sub
