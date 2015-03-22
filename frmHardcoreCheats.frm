VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmHardcoreCheats 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Healing"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5925
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmHardcoreCheats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "arme3"
      Height          =   195
      Left            =   7620
      TabIndex        =   116
      Top             =   4920
      Width           =   735
   End
   Begin JwldButn2b.JeweledButton Command2 
      Height          =   255
      Left            =   4500
      TabIndex        =   115
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "apply"
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
   Begin VB.CheckBox Check2 
      Caption         =   "arme2"
      Height          =   255
      Left            =   7560
      TabIndex        =   114
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3360
      TabIndex        =   113
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   112
      Text            =   "-"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4320
      TabIndex        =   111
      Text            =   "-"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "apply"
      Height          =   255
      Left            =   6780
      TabIndex        =   109
      Top             =   300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3360
      TabIndex        =   103
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3360
      TabIndex        =   102
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3600
      TabIndex        =   101
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3360
      TabIndex        =   100
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3360
      TabIndex        =   99
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8400
      TabIndex        =   93
      Text            =   "0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8400
      TabIndex        =   92
      Text            =   "0"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8400
      TabIndex        =   91
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8400
      TabIndex        =   90
      Text            =   "0"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8400
      TabIndex        =   89
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   -1560
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   -1560
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "arme"
      Height          =   195
      Left            =   7560
      TabIndex        =   87
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Timer TimerPtheal 
      Interval        =   300
      Left            =   4800
      Top             =   2760
   End
   Begin VB.TextBox txtExuraVitaMana 
      Height          =   285
      Left            =   4920
      TabIndex        =   66
      Text            =   "160"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtExuraVitaMana2 
      Height          =   285
      Left            =   4920
      TabIndex        =   72
      Text            =   "70"
      Top             =   720
      Width           =   615
   End
   Begin VB.HScrollBar scrollHP4 
      Height          =   255
      Left            =   8760
      Max             =   100
      TabIndex        =   84
      Top             =   6480
      Value           =   40
      Width           =   1455
   End
   Begin VB.TextBox txtExuraVita4 
      Height          =   285
      Left            =   1140
      TabIndex        =   83
      Text            =   "SELF UHEAL"
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CheckBox chkAutoVita4 
      Caption         =   "Heal :"
      Height          =   255
      Left            =   2520
      TabIndex        =   82
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtExuraVita3 
      Height          =   285
      Left            =   1140
      TabIndex        =   79
      Text            =   "SELF MANA"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.HScrollBar scrollHP3 
      Height          =   255
      Left            =   8760
      Max             =   100
      TabIndex        =   78
      Top             =   6120
      Value           =   70
      Width           =   1455
   End
   Begin VB.CheckBox chkAutoVita3 
      BackColor       =   &H80000018&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   2520
      TabIndex        =   77
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox chkAutoVita2 
      BackColor       =   &H80000018&
      Caption         =   "Spell Hi :"
      Height          =   255
      Left            =   840
      TabIndex        =   75
      Top             =   4080
      Width           =   975
   End
   Begin VB.HScrollBar scrollHP22 
      Height          =   255
      Left            =   8760
      Max             =   100
      TabIndex        =   74
      Top             =   5040
      Value           =   70
      Width           =   1455
   End
   Begin VB.TextBox txtExuraVita2 
      Height          =   285
      Left            =   1140
      TabIndex        =   73
      Text            =   "exura gran"
      Top             =   720
      Width           =   1275
   End
   Begin VB.CheckBox chkAutoVita 
      BackColor       =   &H80000018&
      Caption         =   "Spell Lo :"
      Height          =   255
      Left            =   840
      TabIndex        =   69
      Top             =   4320
      Width           =   1095
   End
   Begin VB.HScrollBar scrollHP2 
      Height          =   255
      Left            =   8760
      Max             =   100
      TabIndex        =   68
      Top             =   5400
      Value           =   70
      Width           =   1455
   End
   Begin VB.TextBox txtExuraVita 
      Height          =   285
      Left            =   1140
      TabIndex        =   67
      Text            =   "exura vita"
      Top             =   1080
      Width           =   1275
   End
   Begin VB.HScrollBar scrollHP 
      Height          =   255
      Left            =   8760
      Max             =   100
      TabIndex        =   26
      Top             =   8280
      Value           =   60
      Width           =   1455
   End
   Begin VB.CheckBox chkAutoHeal 
      BackColor       =   &H80000018&
      Caption         =   "UH Rune"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.TextBox txtExuraVitaMana3 
      Height          =   285
      Left            =   8040
      TabIndex        =   80
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkProtectedShots 
      BackColor       =   &H00000000&
      Caption         =   "Avoid shoting damage runes if your %hp < AutoRuneHeal %hp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   65
      Top             =   7560
      Width           =   5175
   End
   Begin VB.CheckBox chkGmMessagesPauseAll 
      BackColor       =   &H00000000&
      Caption         =   "Gm messages trigger special events and pauses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   64
      Top             =   6240
      Width           =   4215
   End
   Begin VB.Frame frmNewCheats 
      BackColor       =   &H80000014&
      Caption         =   "Heal Method"
      Height          =   975
      Left            =   120
      TabIndex        =   60
      Top             =   2640
      Width           =   5655
      Begin VB.Timer TimerAUH 
         Interval        =   300
         Left            =   4080
         Top             =   120
      End
      Begin VB.Timer TimerPmheal 
         Interval        =   300
         Left            =   5160
         Top             =   120
      End
      Begin VB.Timer TimerSplo 
         Interval        =   100
         Left            =   3600
         Top             =   120
      End
      Begin VB.Timer TimerSphi 
         Interval        =   550
         Left            =   3120
         Top             =   120
      End
      Begin VB.OptionButton chkClassic 
         BackColor       =   &H80000014&
         Caption         =   "Classic mode. for Old Tibia Clients"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   180
         TabIndex        =   63
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton chkEnhancedCheats 
         BackColor       =   &H80000018&
         Caption         =   "No need to open bps, exact cast. Little chance of waste."
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   3660
         TabIndex        =   62
         Top             =   240
         Width           =   15
      End
      Begin VB.OptionButton chkTotalWaste 
         BackColor       =   &H80000014&
         Caption         =   "Hotkeys. leave this default for newer versions"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   180
         TabIndex        =   61
         Top             =   600
         Value           =   -1  'True
         Width           =   3555
      End
   End
   Begin VB.TextBox tibiaTittleFormat 
      Height          =   285
      Left            =   5460
      TabIndex        =   59
      Text            =   "$charactername$ - $expleft$ exp to lv $nextlevel$ - $exph$ exp/h"
      Top             =   8595
      Width           =   3255
   End
   Begin VB.ComboBox cmbWhere 
      Height          =   315
      Left            =   8640
      TabIndex        =   55
      Text            =   "19 : white center"
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox txtExivaExpFormat 
      Height          =   285
      Left            =   6960
      TabIndex        =   53
      Text            =   $"frmHardcoreCheats.frx":014A
      Top             =   7800
      Width           =   3255
   End
   Begin VB.TextBox txtRelogBackpacks 
      Height          =   285
      Left            =   6120
      TabIndex        =   51
      Text            =   "4"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CheckBox chkAutorelog 
      BackColor       =   &H00000000&
      Caption         =   "Auto relog in kick or serversave and reopen"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   50
      Top             =   6840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox txtBlueauraDelay 
      Height          =   285
      Left            =   7920
      TabIndex        =   48
      Text            =   "300"
      Top             =   6840
      Width           =   735
   End
   Begin VB.CheckBox chkAutoGratz 
      BackColor       =   &H00000000&
      Caption         =   "Auto gratz at level advances"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CheckBox chkCaptionExp 
      BackColor       =   &H00000000&
      Caption         =   "Show exp in Tibia window title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   44
      Top             =   6240
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton cmdBigMap 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Big map"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox txtAlarmUHs 
      Height          =   285
      Left            =   5160
      TabIndex        =   42
      Text            =   "5"
      Top             =   6840
      Width           =   495
   End
   Begin VB.CheckBox chkRuneAlarm 
      BackColor       =   &H00000000&
      Caption         =   "When autohealing, alarm when UHS <"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   41
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtRemoteLeader 
      Height          =   285
      Left            =   4920
      TabIndex        =   40
      Top             =   7680
      Width           =   735
   End
   Begin VB.CheckBox chkColorEffects 
      BackColor       =   &H00000000&
      Caption         =   "Show colour effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   38
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox pushID 
      Height          =   375
      Left            =   5040
      TabIndex        =   36
      Text            =   "9"
      Top             =   11160
      Width           =   495
   End
   Begin VB.Timer timerSpam 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   5520
   End
   Begin VB.ComboBox cmbOrderType 
      Height          =   315
      Left            =   2160
      TabIndex        =   33
      Text            =   "type 0 : SD (XYZ)"
      Top             =   8160
      Width           =   2295
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   3240
      TabIndex        =   30
      Text            =   "firenow"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CheckBox chkAcceptSDorder 
      BackColor       =   &H00000000&
      Caption         =   "Accept order if you get in a channel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   8160
      Width           =   3495
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reactivate (will CLOSE proxy)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.HScrollBar scrollLight 
      Height          =   255
      Left            =   2640
      Max             =   15
      TabIndex        =   3
      Top             =   7320
      Value           =   15
      Width           =   1935
   End
   Begin VB.CommandButton cmdOpenBackpacks 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Backpacks"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Map Click action"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3720
      TabIndex        =   21
      Top             =   9120
      Width           =   1815
      Begin VB.OptionButton ActionPath 
         BackColor       =   &H00000000&
         Caption         =   "Move there"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton ActionNothing 
         BackColor       =   &H00000000&
         Caption         =   "Do nothing"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton ActionMove 
         BackColor       =   &H00000000&
         Caption         =   "Summon to bag"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton ActionInspect 
         BackColor       =   &H00000000&
         Caption         =   "Game Inspect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Text            =   "-"
      Top             =   180
      Width           =   1815
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H00000000&
      Caption         =   "ontop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   9720
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkLockOnMyFloor 
      BackColor       =   &H00000000&
      Caption         =   "lock"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   9720
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdateMap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- update now !"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   1575
   End
   Begin VB.OptionButton chkAutoUpdateMap 
      BackColor       =   &H00000000&
      Caption         =   "Full auto update (slow ! )"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   11400
      Width           =   2295
   End
   Begin VB.OptionButton chkUpdateMs 
      BackColor       =   &H00000000&
      Caption         =   "Update each x mseconds:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   11160
      Width           =   2295
   End
   Begin VB.OptionButton chkManualUpdate 
      BackColor       =   &H00000000&
      Caption         =   "No auto update"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   10920
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Timer timerAutoUpdater 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   11520
   End
   Begin VB.TextBox cmdMs 
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Text            =   "1000"
      Top             =   11160
      Width           =   735
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   11160
      Width           =   375
   End
   Begin VB.Timer timerLight 
      Interval        =   1000
      Left            =   4320
      Top             =   5400
   End
   Begin VB.CommandButton cmdOpenTrueRadar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show True Map"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CheckBox chkLight 
      BackColor       =   &H00000000&
      Caption         =   "Change light to this intensity :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkLogoutIfDanger 
      BackColor       =   &H00000000&
      Caption         =   "Logout! if danger on screen at start"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   3135
   End
   Begin VB.CheckBox chkApplyCheats 
      BackColor       =   &H80000018&
      Caption         =   "Activate hardcore cheats  "
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   6000
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.Label Label16 
      Caption         =   "FOR THE CHANGES TAKE EFFECT, HIT THE APPLY BUTTON ON THE BOT"
      Height          =   375
      Left            =   120
      TabIndex        =   110
      Top             =   3645
      Width           =   6240
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000014&
      Caption         =   "UH Rune"
      Height          =   255
      Left            =   300
      TabIndex        =   108
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000014&
      Caption         =   "Mana Pot"
      Height          =   255
      Left            =   300
      TabIndex        =   107
      Top             =   2205
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000014&
      Caption         =   "Heal Pot"
      Height          =   255
      Left            =   300
      TabIndex        =   106
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000014&
      Caption         =   "Spell Lo :"
      Height          =   255
      Left            =   300
      TabIndex        =   105
      Top             =   1125
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000014&
      Caption         =   "Spell Hi :"
      Height          =   255
      Left            =   300
      TabIndex        =   104
      Top             =   765
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000014&
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   98
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   2760
      TabIndex        =   97
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   96
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   95
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   94
      Top             =   780
      Width           =   615
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000002&
      X1              =   5760
      X2              =   5760
      Y1              =   600
      Y2              =   2520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000002&
      X1              =   5760
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000002&
      X1              =   5760
      X2              =   120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   2520
   End
   Begin VB.Label lblHPvalue2 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   10320
      TabIndex        =   70
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblYourPos2 
      BackColor       =   &H80000014&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   4320
      TabIndex        =   71
      Top             =   780
      Width           =   615
   End
   Begin VB.Label lblYourPos 
      BackColor       =   &H80000014&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblHPvalue22 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   10080
      TabIndex        =   76
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label lblHPvalue 
      BackColor       =   &H80000018&
      Caption         =   "60 %"
      Height          =   255
      Left            =   10320
      TabIndex        =   27
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label lblHPvalue4 
      BackColor       =   &H80000018&
      Caption         =   "40 %"
      Height          =   255
      Left            =   10320
      TabIndex        =   86
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label lblHPvalue3 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   10320
      TabIndex        =   85
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label lblYourPos3 
      BackColor       =   &H80000018&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   7440
      TabIndex        =   81
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "tibia window tittle:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   58
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "exiva exp message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   57
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Display exiva exp as:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5460
      TabIndex        =   56
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblRedefineExp 
      BackColor       =   &H00000000&
      Caption         =   "Redefine exp info:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   54
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblBackpacks 
      BackColor       =   &H00000000&
      Caption         =   "bps"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   52
      Top             =   7140
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Blue aura delay between casts ( in mseconds):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   47
      Top             =   7920
      Width           =   3615
   End
   Begin VB.Label lblRedefine 
      BackColor       =   &H00000000&
      Caption         =   "*Redefine ExuraVita:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   46
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9300
      X2              =   8280
      Y1              =   7620
      Y2              =   6240
   End
   Begin VB.Label lblLeader 
      BackColor       =   &H00000000&
      Caption         =   "Only accept order from this leader (leave blank for no leader) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   8520
      Width           =   4455
   End
   Begin VB.Label lblAdvanced 
      BackColor       =   &H00000000&
      Caption         =   "internal p delay"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   35
      Top             =   11160
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   5640
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label lblOn 
      BackColor       =   &H00000000&
      Caption         =   "on targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label lblRead 
      BackColor       =   &H00000000&
      Caption         =   "read order as:  cast"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   32
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label lblOrder2 
      BackColor       =   &H00000000&
      Caption         =   ":targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   31
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00000000&
      Caption         =   "Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   10560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Tile stack for last selected position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   11760
      Width           =   5175
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000014&
      Caption         =   "Auto-Heal :"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lblArraySelected 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   12000
      Width           =   5295
   End
   Begin VB.Label lblLightValue 
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   7200
      Width           =   735
   End
End
Attribute VB_Name = "frmHardcoreCheats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit







Private Sub Check1_Click()
If lock_chkarme = False Then
If HardcoreCheatsIDselected > 0 Then
  If Check1.Value = 1 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme = True
  Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme = False
  End If
End If
End If
End Sub

Private Sub Check2_Click()
If lock_chkarme2 = False Then
If HardcoreCheatsIDselected > 0 Then
  If Check2.Value = 1 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme2 = True
  Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme2 = False
  End If
End If
End If
End Sub

Private Sub Check3_Click()
If lock_chkarme3 = False Then
If HardcoreCheatsIDselected > 0 Then
  If Check3.Value = 1 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme3 = True
  Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme3 = False
  End If
End If
End If
End Sub

Public Sub chkApplyCheats_Click()
  Dim i As Integer
  If chkApplyCheats.Value = 0 Then
    chkRuneAlarm.Value = 0
    chkRuneAlarm.enabled = False
    chkLogoutIfDanger.Value = 0
    chkLight.Value = 0
    frmTrueMap.Hide
    frmBackpacks.Hide
    scrollLight.enabled = False
    chkLogoutIfDanger.enabled = False
    chkLight.enabled = False
    cmdOpenTrueRadar.enabled = False
    lblLightValue.enabled = False
    chkApplyCheats.enabled = False
    chkManualUpdate.Value = True
    chkUpdateMs.Value = False
    chkAutoUpdateMap.Value = False
    chkManualUpdate.enabled = False
    chkUpdateMs.enabled = False
    chkAutoUpdateMap.enabled = False
    timerAutoUpdater.enabled = False
    cmdMs.enabled = False
    cmdChange.enabled = False
    cmdUpdateMap.enabled = False
    Frame1.enabled = False
    chkLockOnMyFloor.enabled = False
    chkOnTop.enabled = False
    cmbCharacter.enabled = False
    ActionInspect.enabled = False
    ActionMove.enabled = False
    ActionNothing.enabled = False
    ActionPath.enabled = False
    cmdOpenBackpacks.enabled = False
    scrollHP.enabled = False
    chkAutoHeal.enabled = False
    chkAutoHeal.Value = 0
    cmdReset.enabled = True
    lblHPvalue.enabled = False
    txtOrder.enabled = False
    lblOrder2.enabled = False
    chkAcceptSDorder.Value = 0
    chkAcceptSDorder.enabled = False
    cmbOrderType.enabled = False
    lblRead.enabled = False
    lblOn.enabled = False
    chkAutoVita.enabled = False
    chkAutoVita.Value = 0
    scrollHP2.enabled = False
    lblHPvalue2.enabled = False
    chkColorEffects.Value = 0
    chkColorEffects.enabled = False
    lblLeader.enabled = False
    txtRemoteLeader.enabled = False
    lblAdvanced.enabled = False
    pushID.enabled = False
    cmdBigMap.enabled = False
    For i = 1 To MAXCLIENTS
      GotPacketWarning(i) = True
    Next i
  End If
End Sub

Private Sub chkAutoHeal_Click()
  Dim i As Integer
  If chkAutoHeal.Value = 0 Then
    For i = 1 To MAXCLIENTS
      RemoveSpamOrder i, 1 'remove  auto UH
    Next i
  End If
End Sub

Private Sub chkAutoUpdateMap_Click()
  timerAutoUpdater.enabled = False
End Sub

Private Sub chkAutoVita_Click()
If lock_chkAutoVita = False Then
If HardcoreCheatsIDselected > 0 Then
  If chkAutoVita.Value = 1 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).splo = True
  Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).splo = False
  End If
End If
End If
End Sub

Private Sub chkAutoVita2_Click()
If lock_chkAutoVita2 = False Then
If HardcoreCheatsIDselected > 0 Then
  If chkAutoVita2.Value = 1 Then
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).sphi = True
  Else
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).sphi = False
  End If
End If
End If
End Sub

Private Sub chkAutoVita3_Click()
If lock_chkAutoVita4 = False Then
If HardcoreCheatsIDselected > 0 Then
  If chkAutoVita4.Value = 1 Then
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).pmh = True
  Else
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).pmh = False
  End If
End If
End If
End Sub

Private Sub chkAutoVita4_Click()
If lock_chkAutoVita4 = False Then
If HardcoreCheatsIDselected > 0 Then
  If chkAutoVita4.Value = 1 Then
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).pth = True
  Else
    'HardcoreCheatsOptions(HardcoreCheatsIDselected).pth = False
  End If
End If
End If
End Sub





Private Sub chkLockOnMyFloor_Click()
  If chkLockOnMyFloor.Value = 1 And cmbCharacter.ListIndex = mapIDselected And mapIDselected > 0 Then
    If mapFloorSelected <> myZ(mapIDselected) Then
      mapFloorSelected = myZ(mapIDselected)
      frmTrueMap.DrawFloor
    End If
  End If
End Sub

Private Sub chkManualUpdate_Click()
  timerAutoUpdater.enabled = False
End Sub

Public Sub chkOnTop_Click()
  If chkOnTop.Value = 1 Then
    ToggleTopmost frmTrueMap.hwnd, True
    ToggleTopmost frmMapReader.hwnd, True
    MapWantedOnTop = True
  Else
    ToggleTopmost frmTrueMap.hwnd, False
    ToggleTopmost frmMapReader.hwnd, False
    MapWantedOnTop = False
  End If
End Sub











Private Sub chkUpdateMs_Click()
  cmdMs_Change
  timerAutoUpdater.enabled = True
End Sub


Private Sub cmbCharacter_Click()
 HardcoreCheatsIDselected = cmbCharacter.ListIndex
  If HardcoreCheatsIDselected > 0 Then
      UpdateValues
  End If
  'mapIDselected = cmbCharacter.ListIndex
  'If mapIDselected > 0 Then
  '    UpdateValues
  'End If
  'If mapIDselected > 0 Then
  '  If TrialVersion = True Then
  '    If GameConnected(mapIDselected) = True And sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
  '        mapFloorSelected = myZ(mapIDselected)
  '        lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
  '        frmTrueMap.SetButtonColours
  '        frmTrueMap.DrawFloor
  '    End If
  '  Else
  '    If GameConnected(mapIDselected) = True Then
  '      mapFloorSelected = myZ(mapIDselected)
  '      lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
  '      frmTrueMap.SetButtonColours
  '      frmTrueMap.DrawFloor
        
   '   End If
   ' End If
  'End If
End Sub









Private Sub cmdBigMap_Click()
  frmMapReader.WindowState = vbNormal
  frmMapReader.Show
  DoEvents
  frmMapReader.ShowCenter
  DoEvents
  frmMapReader.timerBigMapUpdate.enabled = True
End Sub

Public Sub cmdMs_Change()
 Dim lngValue
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lngValue = CLng(cmdMs.Text)
  If lngValue >= 10 And lngValue <= 500000 Then
    timerAutoUpdater.Interval = lngValue
  Else
    cmdMs.Text = "1000"
    timerAutoUpdater.Interval = 1000
  End If
  Exit Sub
gotError:
  cmdMs.Text = "1000"
  timerAutoUpdater.Interval = 1000
End Sub


Private Sub cmdOpenBackpacks_Click()
  frmBackpacks.Show
End Sub

Private Sub cmdOpenTrueRadar_Click()
  frmTrueMap.WindowState = vbNormal
  frmTrueMap.Show
End Sub

Private Sub cmdReset_Click()
  chkApplyCheats.Value = 1
  chkLight.Value = 1
  chkAutoHeal.Value = 1
  frmMenu.Form_Unload False
End Sub

Private Sub cmdUpdateMap_Click()
  If TrialVersion = True Then
    If sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
      frmTrueMap.DrawFloor
    End If
  Else
    frmTrueMap.DrawFloor
  End If
End Sub


Private Sub Combo1_Click()
Dim Index As Integer
Dim idConnection As Integer
If HardcoreCheatsIDselected > 0 Then
If Combo1.ListIndex = 0 Then
'Text1.Text = "77 0D" 'arrow
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self heal"
ElseIf Combo1.ListIndex = 1 Then
'Text1.Text = "79 0D" 'burst
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self sheal"
ElseIf Combo1.ListIndex = 2 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self gheal"
ElseIf Combo1.ListIndex = 3 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self uheal"
ElseIf Combo1.ListIndex = 4 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self pmana"
ElseIf Combo1.ListIndex = 5 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = "self uh"
End If
End If
UpdateValues

End Sub





Private Sub Combo2_Click()
Dim Index As Integer

If HardcoreCheatsIDselected > 0 Then
If Combo2.ListIndex = 0 Then
'Text1.Text = "77 0D" 'arrow
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = "self mana"
ElseIf Combo2.ListIndex = 1 Then
'Text1.Text = "79 0D" 'burst
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = "self smana"
ElseIf Combo2.ListIndex = 2 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = "self gmana"
ElseIf Combo2.ListIndex = 3 Then
'Text1.Text = "C4 1C" 'sniper
HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = "self pmana"
End If
End If
UpdateValues

End Sub

Private Sub Command2_Click()
UpdateValues
End Sub

Private Sub Form_Load()
  LoadHardcoreCheatsChars
  cmbOrderType.Clear
  cmbOrderType.AddItem "type 0 : SD (XYZ)", 0
  cmbOrderType.AddItem "type 1 : HMM (XYZ)", 1
  cmbOrderType.AddItem "type 2 : Explosion (XYZ)", 2
  cmbOrderType.AddItem "type 3 : IH (XYZ)", 3
  cmbOrderType.AddItem "type 4 : UH (XYZ)", 4
  cmbOrderType.AddItem "type 5 : SD (battlelist)", 5
  cmbOrderType.AddItem "type 6 : HMM (battlelist)", 6
  cmbOrderType.AddItem "type 7 : Explosion (battlelist)", 7
  cmbOrderType.AddItem "type 8 : IH (battlelist)", 8
  cmbOrderType.AddItem "type 9 : UH (battlelist)", 9
  cmbOrderType.AddItem "type A : Say (text)", 10
  cmbOrderType.AddItem "type B : fireball (battlelist)", 11
  cmbOrderType.AddItem "type C : stalagmite (battlelist)", 12
  cmbOrderType.AddItem "type D : icicle (battlelist)", 13
  cmbOrderType.Text = "type 5 : SD (battlelist)"
  cmbWhere.Clear
  cmbWhere.AddItem "01 : yellow default"
  cmbWhere.AddItem "02 : yellow default"
  cmbWhere.AddItem "03 : yellow default"
  cmbWhere.AddItem "04 : blue default"
  cmbWhere.AddItem "05 : blue default"
  cmbWhere.AddItem "06 : invisible?"
  cmbWhere.AddItem "07 : invisible?"
  cmbWhere.AddItem "08 : invisible?"
  cmbWhere.AddItem "09 : red default"
  cmbWhere.AddItem "10 : invisible?"
  cmbWhere.AddItem "11 : red default"
  cmbWhere.AddItem "12 : red default"
  cmbWhere.AddItem "13 : red default"
  cmbWhere.AddItem "14 : invisible?"
  cmbWhere.AddItem "15 : red default"
  cmbWhere.AddItem "16 : orange default"
  cmbWhere.AddItem "17 : orange default"
  cmbWhere.AddItem "18 : red center"
  cmbWhere.AddItem "19 : white center"
  cmbWhere.AddItem "20 : 1 line log"
  cmbWhere.AddItem "21 : white system"
  cmbWhere.AddItem "22 : green middle"
  cmbWhere.AddItem "23 : white system"
  cmbWhere.AddItem "24 : purple default"
  cmbWhere.Text = ExivaExpPlace
  Combo1.AddItem "Health Potion", 0
  Combo1.AddItem "Strong Health", 1
  Combo1.AddItem "Great Health", 2
  Combo1.AddItem "Ultimate Health", 3
  Combo1.AddItem "Spirit", 4
  Combo1.AddItem "UH", 5
  Combo2.AddItem "Mana Potion", 0
  Combo2.AddItem "Strong Mana", 1
  Combo2.AddItem "Great Mana", 2
  Combo2.AddItem "Spirit", 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub



Public Sub scrollHP22_Change()
  lblHPvalue22.Caption = CStr(scrollHP22.Value) & " %"
End Sub

Private Sub pushID_Change()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  PUSHDELAYTIMES = CLng(pushID.Text)
  Exit Sub
goterr:
  PUSHDELAYTIMES = 9
  pushID.Text = 9
End Sub

Public Sub scrollHP_Change()
  ChangeGLOBAL_RUNEHEAL_HP scrollHP.Value
End Sub

Public Sub scrollHP2_Change()
  lblHPvalue2.Caption = CStr(scrollHP2.Value) & " %"
End Sub


Private Sub scrollHP3_Change()
  lblHPvalue3.Caption = CStr(scrollHP3.Value) & " %"
End Sub

Private Sub scrollHP4_Change()
  lblHPvalue4.Caption = CStr(scrollHP4.Value) & " %"
End Sub

Public Sub scrollLight_Change()
 lblLightValue.Caption = CStr(Round((scrollLight.Value / 15) * 100)) & " %"
  LightIntesityHex = GoodHex(CByte(scrollLight.Value))
End Sub


Private Sub Text10_Validate(Cancel As Boolean)
Dim idConnection As Integer
'Dim lonN As Long
Dim lngValue As Long
'lngValue = CLng(Text11.Text)

If IsNumeric(Text10.Text) Then
    Text10.Text = Text10.Text
    Text3.Text = Text10.Text
Else
    Text10.Text = "0"
End If

For idConnection = 1 To MAXCLIENTS

If GameConnected(idConnection) = True Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10 = Text10
If IsNumeric(Text10.Text) Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10
Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3
End If
End If

Next idConnection


End Sub

Private Sub Text11_Validate(Cancel As Boolean)
Dim idConnection As Integer
'Dim lonN As Long
Dim lngValue As Long
'lngValue = CLng(Text11.Text)

If IsNumeric(Text11.Text) Then
    Text11.Text = Text11.Text
    Text2.Text = Text11.Text
Else
    Text11.Text = Text2.Text
End If

For idConnection = 1 To MAXCLIENTS

'If GameConnected(idConnection) = True Then
If HardcoreCheatsIDselected > 0 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11 = Text11
If IsNumeric(HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11) Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11
    UpdateValues
Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text2
End If
End If

Next idConnection

Exit Sub
End Sub











Private Sub Text12_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If HardcoreCheatsIDselected > 0 Then
  lonN = CLng(Text12.Text)
  If lonN > 0 Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12 = lonN
  Else
    Text12.Text = CStr(HardcoreCheatsOptions_Text12_default)
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12 = HardcoreCheatsOptions_Text12_default
  End If
  End If
  Exit Sub
gotError:
  Text12.Text = CStr(HardcoreCheatsOptions_Text12_default)
  HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12 = HardcoreCheatsOptions_Text12_default
End Sub

Private Sub Text2_Change()
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).Text2 = Text2.Text
End If
End Sub

Private Sub Text3_Change()
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3 = Text3.Text
End If
End Sub

Private Sub Text5_Change()
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).Text5 = Text5.Text
End If
End Sub

Private Sub Text6_Change()
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).Text6 = Text6.Text
End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Dim idConnection As Integer
'Dim lonN As Long
Dim lngValue As Long
'lngValue = CLng(Text11.Text)

If IsNumeric(Text7.Text) Then
    Text7.Text = Text7.Text
    Text6.Text = Text7.Text
Else
    Text7.Text = "0"
End If

For idConnection = 1 To MAXCLIENTS

If GameConnected(idConnection) = True Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7 = Text7
If IsNumeric(Text7.Text) Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text6 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7
Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text6
End If
End If

Next idConnection


End Sub

Private Sub Text8_Validate(Cancel As Boolean)
Dim idConnection As Integer
'Dim lonN As Long
Dim lngValue As Long
'lngValue = CLng(Text11.Text)

If IsNumeric(Text8.Text) Then
    Text8.Text = Text8.Text
    Text5.Text = Text8.Text
Else
    Text8.Text = "0"
End If

For idConnection = 1 To MAXCLIENTS

If GameConnected(idConnection) = True Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8 = Text8
If IsNumeric(Text8.Text) Then
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text5 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8
Else
    HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8 = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text5
End If
End If

Next idConnection


End Sub

Private Sub Text9_Validate(Cancel As Boolean)
Dim idConnection As Integer
'Dim lonN As Long
Dim lngValue As Long
'lngValue = CLng(Text11.Text)

If IsNumeric(Text9.Text) Then
    Text9.Text = Text9.Text
    Text4.Text = Text9.Text
Else
    Text9.Text = "0"
End If

For idConnection = 1 To MAXCLIENTS

If GameConnected(idConnection) = True Then
If IsNumeric(Text9.Text) Then
    Text9.Text = Text9.Text
    Text4.Text = Text9.Text
Else
    Text9.Text = Text4.Text
End If
End If

Next idConnection


End Sub

Private Sub TimerAUH_Timer()
 Dim idConnection As Integer
  Dim learnResult As TypeLearnResult
  Dim aRes As Long
  
For idConnection = 1 To MAXCLIENTS
If GameConnected(idConnection) = True Then
   If (myHP(idConnection) < CLng(HardcoreCheatsOptions(idConnection).Text12)) And _
             (sentFirstPacket(idConnection) = True) Then
                HardcoreCheatsOptions(idConnection).arme3 = True
                aRes = ExecuteInTibia("self uh", idConnection, True)
                DoEvents
                Else
                HardcoreCheatsOptions(idConnection).arme3 = False
   End If
End If
Next idConnection
End Sub

Private Sub timerAutoUpdater_Timer()
 If mapIDselected > 0 Then
    If TrialVersion = True Then
      If sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
        If chkLockOnMyFloor.Value = 1 Then
          mapFloorSelected = myZ(mapIDselected)
        End If
        frmTrueMap.SetButtonColours
        frmTrueMap.DrawFloor
      End If
    Else
      If chkLockOnMyFloor.Value = 1 Then
        mapFloorSelected = myZ(mapIDselected)
      End If
      frmTrueMap.SetButtonColours
      frmTrueMap.DrawFloor
    End If
  End If
End Sub

Private Sub timerLight_Timer()
  Dim i As Integer
  'Dim cPacket() As Byte
  Dim errorD As Integer
  Dim inRes As Integer
  Dim aRes As Long
  Dim playerS As String
  'Exit Sub '
  #If FinalMode Then
  On Error GoTo endT
  #End If
  If (TrialVersion = False) And (trialSafety4 <> 4) Then
    End
  End If
  If chkApplyCheats.Value = 1 Then
  
  If (Me.chkCaptionExp.Value = 1) Then
    UpdateTibiaTitles
  End If
  
  For i = 1 To HighestConnectionID
      errorD = i
  If (GameConnected(i) = True) And (sentWelcome(i) = True) And (GotPacketWarning(i) = False) Then
    If ReconnectionStage(i) = 3 Then
      If frmBackpacks.totalbpsOpen(i) < CLng(txtRelogBackpacks.Text) Then
        aRes = openBP(i)
        If aRes = -1 Then
          
          ReconnectionStage(i) = 0 'forced
          If frmBackpacks.totalbpsOpen(i) = 0 Then
             If TibiaVersionLong < 790 Then
               ReconnectionStage(i) = 10
               frmMain.DoCloseActions i
               frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "WARNING: Character " & CharacterName(i) & " had no backpack after relog so it was left closed"
               ReconnectionStage(i) = 0
               logoutAllowed(i) = 0
             Else
               aRes = SendLogSystemMessageToClient(i, "It was not possible to open the desired number of backpacks.")
               DoEvents
             End If
          Else
            aRes = SendLogSystemMessageToClient(i, "It was not possible to open the desired number of backpacks.")
            DoEvents
          End If
        Else
          DoEvents
        End If
      Else
        ReconnectionStage(i) = 0
        logoutAllowed(i) = 0
        aRes = SendLogSystemMessageToClient(i, "Successfully opened " & CStr(txtRelogBackpacks.Text) & " containers.")
        DoEvents
      End If
    End If
      ' ALIVE? (45seconds without packet is not good)
    If (lastPing(i) < (GetTickCount() - MaxTimeWithoutServerPackets)) Then
      If frmHardcoreCheats.chkAutorelog.Value = 1 Then
        aRes = GiveGMmessage(i, "ISP - server down detected (too much time without receiving anything from server)", "Blackdproxy")
        DoEvents
        lastPing(i) = GetTickCount() + 3600000
        StartReconnection i
      Else
        aRes = GiveGMmessage(i, "ISP - server down detected (too much time without receiving anything from server)", "Blackdproxy")
        DoEvents
        lastPing(i) = GetTickCount()
        If frmRunemaker.chkCloseSound.Value = 1 Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "(Giving alarm because server or ISP went probably down)"
          'ChangePlayTheDangerSound True
        End If
      End If
    End If
    If CheatsPaused(i) = False Then
      If (RuneMakerOptions(i).msgSound2 = True) Then
        playerS = PlayerOnScreen(i)
        If playerS <> "" Then
        'If DangerPlayer(i) = True Then
        '  playerS = DangerPlayerName(i)
          PlayMsgSound2 = True
          If publicDebugMode = True Then
            aRes = SendLogSystemMessageToClient(i, "[Debug] Giving alarm because you have on screen: " & playerS)
            DoEvents
          End If
        End If
      End If

      If DangerGM(i) = True Then
        If (GetTickCount() > LogoutTimeGM(i)) And (LogoutTimeGM(i) <> 0) Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(i) & " closed after random time,because : " & GMname(i)
          aRes = GiveServerError("Closed after random time, because : " & GMname(i), i)
          DoEvents
          ReconnectionStage(i) = 10
          frmMain.DoCloseActions i
          ReconnectionStage(i) = 0
          DoEvents
          Exit Sub
        End If
       ' If frmRunemaker.ChkDangerSound.Value = 1 Then
          'ChangePlayTheDangerSound True
        'End If
      End If
      If DangerPK(i) = True Then
          'If PlayTheDangerSound = False Then
            'If frmCavebot.chkChangePkHeal.Value = 1 Then
              'ChangeGLOBAL_RUNEHEAL_HP frmCavebot.scrollPkHeal.Value
            'Else

            'End If
            'aRes = SendLogSystemMessageToClient(i, "Blackd Proxy NG: To deactivate alarm say 'exiva cancel'")
            'DoEvents
          'End If
          'ChangePlayTheDangerSound True
      End If
    End If
      If (chkLight.Value = 1) Then
        If IDstring(i) <> "" Then
          If frmMain.sckClientGame(i).State = sckConnected Then
            If PlayTheDangerSound = True Then
              'nextLight(i) = "FD"
              'enLight i
              'nextLight(i) = "D7"
            ElseIf nextLight(i) <> "D7" Then
              'nextLight(i) = "D7"
            Else
              'enLight i
            End If
          Else
            'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client #" & CStr(i) & "# lost connection during timerLight_Timer"
            'frmMain.DoCloseActions i
            'DoEvents
          End If
        End If
      End If
  End If
  Next i
  

  
  If PlayMsgSound = True Then
    'PlayMsgSound = False
    If frmRunemaker.ChkDangerSound.Value = 1 Then
        'DirectX_PlaySound 4 ' play ding.wav
    End If
    If ((frmRunemaker.chkOnDangerSS.Value = 1) And (frmRunemaker.timerSS.enabled = False)) Then
        frmRunemaker.timerSS.enabled = True
    End If
  End If
  If PlayMsgSound2 = True Then
    PlayMsgSound2 = False
    If frmRunemaker.ChkDangerSound.Value = 1 Then
        'DirectX_PlaySound 1 ' play player.wav
    End If
    If ((frmRunemaker.chkOnDangerSS.Value = 1) And (frmRunemaker.timerSS.enabled = False)) Then
        frmRunemaker.timerSS.enabled = True
    End If
  End If
  If (PlayTheDangerSound = True) Then ' And (frmRunemaker.ChkDangerSound.Value = 1) Then
    If frmRunemaker.ChkDangerSound.Value = 1 Then
       ' DirectX_PlaySound 2 ' play danger.wav
    End If
  End If
      
  End If
  Exit Sub
endT:
  On Error GoTo severeE:
  If PlayMsgSound = True Then
    PlayMsgSound = False
    'DirectX_PlaySound 3 ' play ding.wav
  End If
  If PlayMsgSound2 = True Then
    PlayMsgSound2 = False
    'DirectX_PlaySound 1 ' play player.wav
  End If
  If (PlayTheDangerSound = True) And (frmRunemaker.ChkDangerSound.Value = 1) Then
    'DirectX_PlaySound 2 ' play danger.wav
  End If
severeE:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got unexpected error at timerLight_Timer - Client ID terminated!"
  frmMain.DoCloseActions i
  DoEvents
End Sub

Private Sub TimerPmheal_Timer()
  Dim idConnection As Integer
  Dim learnResult As TypeLearnResult
  Dim aRes As Long
  TimerPmheal.enabled = True
  
For idConnection = 1 To MAXCLIENTS
If GameConnected(idConnection) = True Then
   If (myMana(idConnection) < CLng(HardcoreCheatsOptions(idConnection).Text5)) And _
             (sentFirstPacket(idConnection) = True) And (HardcoreCheatsOptions(idConnection).arme2 = False) And (HardcoreCheatsOptions(idConnection).arme3 = False) Then
                aRes = ExecuteInTibia(HardcoreCheatsOptions(idConnection).txtExuraVita3, idConnection, True)
                DoEvents
   End If
End If
Next idConnection
End Sub

Private Sub TimerPtheal_Timer()
  Dim idConnection As Integer
  Dim learnResult As TypeLearnResult
  Dim aRes As Long
  
For idConnection = 1 To MAXCLIENTS
If GameConnected(idConnection) = True Then
   If (myHP(idConnection) < CLng(HardcoreCheatsOptions(idConnection).Text6)) And _
             (sentFirstPacket(idConnection) = True) Then
                HardcoreCheatsOptions(idConnection).arme2 = True
                aRes = ExecuteInTibia(HardcoreCheatsOptions(idConnection).txtExuraVita4, idConnection, True)
                DoEvents
    Else
                HardcoreCheatsOptions(idConnection).arme2 = False
   End If
End If
Next idConnection

End Sub

Private Sub timerSpam_Timer()
  Dim i As Integer
  Dim order As Integer
  Dim cid As Integer
  Dim resA As Long
  Dim gtc As Long
  Dim posSetting As Long
  Dim act As String
  #If FinalMode Then
  On Error GoTo errIgnore
  #End If
  For i = 1 To MAXCLIENTS
    If (GameConnected(i) = True) Then
      If ((SpamAutoFastHeal(i) = True) And _
       (CheatsPaused(i) = False)) Then
        gtc = GetTickCount()
        If (nextFastHeal(i) <= gtc) Then
          resA = UseFastUH(i)
          If resA = 0 Then
            cancelAllMove(i) = GetTickCount() + 500
            If frmHardcoreCheats.chkColorEffects.Value = 1 Then
              If Not (nextLight(i) = "04") Then
                nextLight(i) = "04"
               enLight i
             End If
            End If
          End If
          nextFastHeal(i) = gtc + BlueAuraDelay
        End If
       ElseIf (SpamAutoHeal(i) = True) And _
      ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
        UHRetryCount(i) = UHRetryCount(i) + 1
          If (UHRetryCount(i) < 50) Then
            If (TibiaVersionLong < 780) Then
                cancelAllMove(i) = GetTickCount() + 500
            End If
          ElseIf (PlayTheDangerSound = False) Then
            'give msg !
            If chkClassic.Value = False Then
                resA = GiveGMmessage(i, "No UH", "Warning")
                ChangePlayTheDangerSound False
                DoEvents
            End If
          End If
          'heal
          resA = UseUH(i)
          If resA = 0 Then
            If frmHardcoreCheats.chkColorEffects.Value = 1 Then
              If Not (nextLight(i) = "04") Then
                nextLight(i) = "04"
                enLight i
              End If
            End If
          End If
      ElseIf ((SpamAutoPush(i) = True)) Then
        If pushDelay(i) = 0 Then
          resA = DoPush(i)
          pushDelay(i) = PUSHDELAYTIMES
          DoEvents
        Else
          pushDelay(i) = pushDelay(i) - 1
        End If
      ElseIf ((SpamAutoMana(i) = True) And _
       (CheatsPaused(i) = False)) Then
       ' Note: it can be a problem when there are no manas left!
        resA = UseFluid(i, byteMana)
      End If
    End If
  Next i
  Exit Sub
errIgnore:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unexpected error at timerSpam_Timer()"
End Sub

Private Function GetOneTitle(idConnection As Integer) As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  'UpdateExpVars idConnection
  var_lf(idConnection) = ". "
  GetOneTitle = parseVars(idConnection, tibiaTittleFormat.Text)
  Exit Function
goterr:
  GetOneTitle = "Tibia"
End Function

Private Sub UpdateTibiaTitles()
  Dim i As Integer
  Dim Message As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If frmStealth.chkStealthExp = 1 Then
    If stealthIDselected <> 0 Then
        If GameConnected(stealthIDselected) = True Then
            Message = GetOneTitle(stealthIDselected)
            frmStealth.Caption = Message
        End If
    End If
  Else
  GetProcessAllProcessIDs
  For i = 1 To MAXCLIENTS
    If (GameConnected(i) = True) Then
      Message = GetOneTitle(i)
      If ProcessID(i) > 0 Then
        SetWindowText ProcessID(i), Message
      End If
    End If
  Next i
  End If
goterr:
  ' just end...
End Sub

Private Sub TimerSphi_Timer()
  Dim idConnection As Integer
  Dim learnResult As TypeLearnResult
  Dim aRes As Long
  'TimerSphi.enabled = True
  
For idConnection = 1 To MAXCLIENTS

If GameConnected(idConnection) = True Then
   If (myHP(idConnection) < CLng(HardcoreCheatsOptions(idConnection).Text2)) And (myMana(idConnection) >= CLng(HardcoreCheatsOptions(idConnection).txtExuraVitaMana2)) And (HardcoreCheatsOptions(idConnection).arme = False) Then
             '(HardcoreCheatsOptions(HardcoreCheatsIDselected).sphi = True) And
                aRes = ExecuteInTibia(HardcoreCheatsOptions(idConnection).txtExuraVita2, idConnection, True)
                DoEvents
   End If
End If

Next idConnection
End Sub

Private Sub TimerSplo_Timer()
  Dim idConnection As Integer
  Dim learnResult As TypeLearnResult
  Dim aRes As Long
  
For idConnection = 1 To MAXCLIENTS
If GameConnected(idConnection) = True Then
   If (myHP(idConnection) < CLng(HardcoreCheatsOptions(idConnection).Text3)) And (sentFirstPacket(idConnection) = True) And (myMana(idConnection) >= CLng(HardcoreCheatsOptions(idConnection).txtExuraVitaMana)) Then
             '(HardcoreCheatsOptions(HardcoreCheatsIDselected).splo = True) And
                HardcoreCheatsOptions(idConnection).arme = True
                aRes = ExecuteInTibia(HardcoreCheatsOptions(idConnection).txtExuraVita, idConnection, True)
                DoEvents
    Else
                HardcoreCheatsOptions(idConnection).arme = False
   End If
End If
Next idConnection
End Sub



Private Sub txtBlueauraDelay_Change()
 Dim lngValue
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lngValue = CLng(txtBlueauraDelay.Text)
  If lngValue >= 10 And lngValue <= 500000 Then
    BlueAuraDelay = lngValue
  Else
    txtBlueauraDelay.Text = "200"
    BlueAuraDelay = 200
  End If
  Exit Sub
gotError:
  txtBlueauraDelay.Text = "200"
  BlueAuraDelay = 200
End Sub


Public Sub UpdateValues()
Dim i As Integer
Dim idConnection As Integer

   If HardcoreCheatsIDselected = 0 Then
    frmHardcoreCheats.Text11.Text = HardcoreCheatsOptions_Text11_default
    frmHardcoreCheats.Text10.Text = HardcoreCheatsOptions_Text10_default
    frmHardcoreCheats.Text7.Text = HardcoreCheatsOptions_Text7_default
    frmHardcoreCheats.Text8.Text = HardcoreCheatsOptions_Text8_default
    frmHardcoreCheats.Text2.Text = HardcoreCheatsOptions_Text2_default
    frmHardcoreCheats.Text12.Text = HardcoreCheatsOptions_Text12_default
    frmHardcoreCheats.Text3.Text = HardcoreCheatsOptions_Text3_default
    frmHardcoreCheats.Text6.Text = HardcoreCheatsOptions_Text6_default
    frmHardcoreCheats.Text5.Text = HardcoreCheatsOptions_Text5_default
    frmHardcoreCheats.txtExuraVita2.Text = HardcoreCheatsOptions_txtExuraVita2_default
    frmHardcoreCheats.txtExuraVitaMana2.Text = HardcoreCheatsOptions_txtExuraVitaMana2_default
    frmHardcoreCheats.txtExuraVitaMana.Text = HardcoreCheatsOptions_txtExuraVitaMana_default
    frmHardcoreCheats.txtExuraVita4.Text = HardcoreCheatsOptions_txtExuraVita4_default
    frmHardcoreCheats.txtExuraVita3.Text = HardcoreCheatsOptions_txtExuraVita3_default
    frmHardcoreCheats.txtExuraVita.Text = HardcoreCheatsOptions_txtExuraVita_default
  Else
  'aqui a validacao
    frmHardcoreCheats.Text6.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7
    frmHardcoreCheats.Text5.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8
    frmHardcoreCheats.Text3.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10
    frmHardcoreCheats.Text2.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11
    frmHardcoreCheats.Text12.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12
    
    frmHardcoreCheats.Text11.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11
    frmHardcoreCheats.Text10.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10
    frmHardcoreCheats.Text7.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7
    frmHardcoreCheats.Text8.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8
    frmHardcoreCheats.Text2.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text2
    frmHardcoreCheats.Text12.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12
    frmHardcoreCheats.Text3.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3
    frmHardcoreCheats.Text6.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text6
    frmHardcoreCheats.Text5.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).Text5
    frmHardcoreCheats.txtExuraVita3.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3
    frmHardcoreCheats.txtExuraVita4.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4
    frmHardcoreCheats.txtExuraVita2.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita2
    frmHardcoreCheats.txtExuraVitaMana2.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana2
    frmHardcoreCheats.txtExuraVitaMana.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana
    frmHardcoreCheats.txtExuraVita.Text = HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita
  End If

  If mapIDselected = 0 Then
    If chkAutoVita2 = True Then
      chkAutoVita2.Value = 1
    Else
      chkAutoVita2.Value = 0
    End If
    If chkAutoVita = True Then
      chkAutoVita.Value = 1
    Else
      chkAutoVita.Value = 0
    End If
    If chkAutoHeal = True Then
      chkAutoHeal.Value = 1
    Else
      chkAutoHeal.Value = 0
    End If
    If chkAutoVita3 = True Then
      chkAutoVita3.Value = 1
    Else
      chkAutoVita3.Value = 0
    End If
  Else
  End If
End Sub

Private Sub txtExuraVita_Validate(Cancel As Boolean)
'splo
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita = txtExuraVita.Text
End If
End Sub

Private Sub txtExuraVita2_Validate(Cancel As Boolean)
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita2 = txtExuraVita2.Text
End If
End Sub

Private Sub txtExuraVita3_Validate(Cancel As Boolean)
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = txtExuraVita3.Text
End If
End Sub

Private Sub txtExuraVita4_Validate(Cancel As Boolean)
If HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = txtExuraVita4.Text
End If
End Sub

Public Sub LoadHardcoreCheatsChars()
  Dim i As Long
  Dim firstC As Long
  If TibiaVersionLong >= 872 Then
   ' UseRightHand.Visible = False
   ' UseLeftHand.Visible = False
   ' fraNoHands.Visible = True
   ' Label2.Caption = "IMPORTANT: Blackd Proxy will only count runes displayed in opened backpacks!"
  End If
  firstC = 0
  cmbCharacter.Clear
  cmbCharacter.AddItem "-", 0
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True Then
      If firstC = 0 Then
        firstC = i
      End If
      cmbCharacter.AddItem CharacterName(i), i
    Else
      cmbCharacter.AddItem "-" & CStr(i) & "- NOT CONNECTED", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  HardcoreCheatsIDselected = firstC
  UpdateValues
End Sub


Private Sub txtExuraVitaMana_Validate(Cancel As Boolean)
If IsNumeric(txtExuraVitaMana.Text) And HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana = txtExuraVitaMana.Text
End If
End Sub

Private Sub txtExuraVitaMana2_Validate(Cancel As Boolean)
If IsNumeric(txtExuraVitaMana2.Text) And HardcoreCheatsIDselected > 0 Then
  HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana2 = txtExuraVitaMana2.Text
End If
End Sub
