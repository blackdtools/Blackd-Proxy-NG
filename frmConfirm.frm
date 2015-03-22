VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmConfirm 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Close"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton cmdNo 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "No"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton cmdYes 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "Yes"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.CommandButton cmdNo2 
      BackColor       =   &H80000018&
      Caption         =   "No"
      Height          =   375
      Left            =   1620
      MaskColor       =   &H80000008&
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes2 
      BackColor       =   &H80000018&
      Caption         =   "Yes"
      Height          =   375
      Left            =   360
      MaskColor       =   &H80000007&
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H80000014&
      Caption         =   "Close Blackd Proxy NG ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   3375
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub cmdNo_Click()
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    confirmedExit = True
    'Unload frmMenu
    Unload frmConfirm
    End
End Sub




Private Sub Form_Unload(Cancel As Integer)
'  Me.Hide
'  Cancel = BlockUnload
'End Sub

' Unload all
  If thisShouldNotBeLoading = 0 Then
    Exit Sub
  End If
  If confirmedExit = False Then
    Cancel = True
    ToggleTopmost frmConfirm.hwnd, True
    frmConfirm.Show
    frmConfirm.SetFocus
    Exit Sub
  End If
  #If DoSave Then
    If LoadWasCompleted = True Then ' check to avoid ini corruption if there was an unexpected fail at loading
      frmMain.WriteIni
    End If
  #End If
  BlockUnload = 0
  frmMain.Timer1.enabled = False 'should not be needed ... just in case
  frmEvents.timerScheduledActions.enabled = False
  'frmTrainer.timerTrainer.enabled = False
  frmCondEvents.timerCheck.enabled = False
 'this removes the icon from the system tray
  Shell_NotifyIcon NIM_DELETE, nid
  Refresh 'ensure icon is deleted from tray
  'LogOnFile "debug.txt", "Ended by user"
  Unload frmMain
  Unload frmCheats
  Unload frmBigText
  Unload frmCavebot
  Unload frmTrueMap
  Unload frmBackpacks
  Unload frmRunemaker
  Unload frmHardcoreCheats
  Unload frmAdvanced
  Unload frmHotkeys
  Unload frmMagebomb
  Unload frmScreenshot
  Unload frmCondEvents
  Unload frmHPmana
  Unload frmNews
  Unload frmStealth
  Unload frmBroadcast
  Set DirectX = Nothing
  Set DX = Nothing
 
  End 'ends all subthreads of this process
  
  End Sub
