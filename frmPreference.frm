VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPreference 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preference"
   ClientHeight    =   4635
   ClientLeft      =   2235
   ClientTop       =   2415
   ClientWidth     =   6645
   Icon            =   "frmPreference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   2700
      TabIndex        =   1
      Top             =   0
      Width           =   3840
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4065
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7170
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPreference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*´°ÌåÃû³Æ£ºfrmPreference
'*´°Ìå¹¦ÄÜ£ºÈí¼þÅäÖÃ
'*ËµÃ÷£ºÈç¹û×öÁË¸Ä¶¯£¬ÇëÓÊ¸øÎÒÒ»·Ý(progame@cnnb.net)
'*      Ö÷Ò³£ºhttp://www.nettoolx.com/progame
'*×÷Õß£ºprogame  2002-10-07  17:56:22
'***************************************************************
Option Explicit

