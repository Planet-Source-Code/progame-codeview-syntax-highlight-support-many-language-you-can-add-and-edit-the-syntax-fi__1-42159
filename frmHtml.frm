VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHtml 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert to Html"
   ClientHeight    =   3945
   ClientLeft      =   2205
   ClientTop       =   1920
   ClientWidth     =   6075
   Icon            =   "frmHtml.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtHtml 
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   6297
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmHtml.frx":000C
   End
End
Attribute VB_Name = "frmHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*´°ÌåÃû³Æ£ºfrmHtml
'*´°Ìå¹¦ÄÜ£ºÏÔÊ¾HTMLºÍUBB¸ñÊ½
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-10-18  9:21:16
'***************************************************************
Option Explicit

