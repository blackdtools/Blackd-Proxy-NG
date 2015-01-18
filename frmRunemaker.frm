VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmRunemaker 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extras"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   Icon            =   "frmRunemaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkautoPM2 
      Caption         =   "Msg that contains :"
      Height          =   195
      Left            =   2400
      TabIndex        =   93
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtbeep 
      Height          =   285
      Left            =   4200
      TabIndex        =   92
      Text            =   "are you there?"
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Click Reuse"
      Height          =   195
      Left            =   240
      TabIndex        =   91
      Top             =   3720
      Value           =   2  'Grayed
      Width           =   1275
   End
   Begin VB.Timer TimerHur2 
      Interval        =   200
      Left            =   7200
      Top             =   1920
   End
   Begin VB.CheckBox chkautoarme6 
      Caption         =   "arme6"
      Height          =   255
      Left            =   6180
      TabIndex        =   90
      Top             =   5220
      Width           =   1035
   End
   Begin VB.CheckBox chkautoarme5 
      Caption         =   "arme5"
      Height          =   255
      Left            =   6180
      TabIndex        =   89
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CheckBox chkautoda 
      Caption         =   "desired action"
      Height          =   195
      Left            =   6180
      TabIndex        =   88
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CheckBox chkautoarme4 
      Caption         =   "arme4"
      Height          =   195
      Left            =   6180
      TabIndex        =   87
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox chkautora 
      Caption         =   "retry atk"
      Height          =   195
      Left            =   6180
      TabIndex        =   86
      Top             =   4140
      Width           =   1035
   End
   Begin VB.CheckBox chkautoee 
      Caption         =   "auto explo"
      Height          =   255
      Left            =   6180
      TabIndex        =   85
      Top             =   3840
      Width           =   1095
   End
   Begin JwldButn2b.JeweledButton cmdApply 
      Height          =   255
      Left            =   4500
      TabIndex        =   84
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
   Begin VB.Timer Timerxray 
      Interval        =   4000
      Left            =   1740
      Top             =   600
   End
   Begin VB.CheckBox chkautoxray 
      Caption         =   "Xray view"
      Height          =   195
      Left            =   240
      TabIndex        =   82
      Top             =   3420
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "MW timer"
      Height          =   195
      Left            =   240
      TabIndex        =   81
      Top             =   3120
      Value           =   2  'Grayed
      Width           =   1335
   End
   Begin VB.Timer Timertestera2 
      Interval        =   100
      Left            =   3720
      Top             =   2340
   End
   Begin VB.CheckBox chkautodk 
      Caption         =   "Diagonal keys"
      Height          =   195
      Left            =   240
      TabIndex        =   79
      Top             =   2820
      Width           =   1455
   End
   Begin VB.CheckBox chkautodd 
      Caption         =   "Dash"
      Height          =   195
      Left            =   240
      TabIndex        =   78
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer TimerPM2 
      Interval        =   100
      Left            =   5340
      Top             =   3360
   End
   Begin VB.Timer Timererg 
      Interval        =   200
      Left            =   6720
      Top             =   3240
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   76
      Text            =   "0"
      Top             =   2820
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   75
      Text            =   "0"
      Top             =   2820
      Width           =   615
   End
   Begin VB.CheckBox chkerg 
      Caption         =   "Energy ring :"
      Height          =   255
      Left            =   2400
      TabIndex        =   72
      Top             =   2460
      Width           =   1335
   End
   Begin VB.Timer Timerssap 
      Interval        =   300
      Left            =   6720
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   71
      Text            =   "50"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox chkEnableTrainer 
      BackColor       =   &H80000018&
      Caption         =   "Active collectitems"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   67
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H80000014&
      Caption         =   "Refill Arrow slot"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   66
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtPickupID 
      Height          =   285
      Left            =   4200
      TabIndex        =   65
      Text            =   "00 00"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   10
      Left            =   3240
      TabIndex        =   64
      Text            =   "00 00"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Items ID (hex)"
      Height          =   315
      Left            =   4980
      TabIndex        =   63
      Top             =   5460
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRunemaker.frx":014A
      Left            =   4080
      List            =   "frmRunemaker.frx":014C
      TabIndex        =   62
      Text            =   "-"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Timer Timerssa 
      Interval        =   200
      Left            =   8760
      Top             =   2520
   End
   Begin VB.CheckBox chkautossa 
      Caption         =   "ssa"
      Height          =   255
      Left            =   8760
      TabIndex        =   61
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timerpmax 
      Interval        =   200
      Left            =   7920
      Top             =   1320
   End
   Begin VB.CheckBox chkautopmax 
      Caption         =   "pushmax"
      Height          =   255
      Left            =   6480
      TabIndex        =   60
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "more"
      Height          =   375
      Left            =   5160
      TabIndex        =   59
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timertar 
      Interval        =   200
      Left            =   6240
      Top             =   960
   End
   Begin VB.CheckBox chkautotar 
      Caption         =   "Hold Attack"
      Height          =   255
      Left            =   6240
      TabIndex        =   58
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtDashDelay 
      Height          =   315
      Left            =   7800
      TabIndex        =   56
      Text            =   "0"
      Top             =   2940
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkDash 
      Caption         =   "Dash"
      Height          =   255
      Left            =   7320
      TabIndex        =   55
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer TimerAp 
      Interval        =   300
      Left            =   9240
      Top             =   1920
   End
   Begin VB.CheckBox chkautoAp 
      Caption         =   "anti push"
      Height          =   255
      Left            =   7680
      TabIndex        =   54
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer TimerSdt 
      Interval        =   200
      Left            =   8760
      Top             =   1920
   End
   Begin VB.CheckBox chkautoSdt 
      Caption         =   "sd target"
      Height          =   255
      Left            =   7680
      TabIndex        =   53
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer TimerDan2 
      Interval        =   3000
      Left            =   8280
      Top             =   1920
   End
   Begin VB.CheckBox chkautoDan 
      Caption         =   "Anti idle"
      Height          =   195
      Left            =   240
      TabIndex        =   52
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer TimergHur2 
      Interval        =   200
      Left            =   7800
      Top             =   1920
   End
   Begin VB.CheckBox chkautogHur 
      Caption         =   "Auto ghaste"
      Height          =   195
      Left            =   240
      TabIndex        =   51
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CheckBox chkautoHur 
      Caption         =   "Auto haste"
      Height          =   195
      Left            =   240
      TabIndex        =   50
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer TimerUtamo 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7320
      Top             =   1320
   End
   Begin VB.CheckBox chkautoUtamo 
      Caption         =   "Mana shield"
      Height          =   195
      Left            =   240
      TabIndex        =   49
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CheckBox chkReveal 
      BackColor       =   &H80000018&
      Caption         =   "Reveal invisible"
      Height          =   255
      Left            =   5040
      TabIndex        =   48
      Top             =   6360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraNoHands 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Tibia 8.72+"
      ForeColor       =   &H00FFFFFF&
      Height          =   535
      Left            =   8880
      TabIndex        =   46
      Top             =   6720
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton NoHands 
         BackColor       =   &H00000000&
         Caption         =   "Don't move runes to hands"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.TextBox txrRunemakerChaos2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9420
      TabIndex        =   44
      Text            =   "1"
      Top             =   660
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveRunemakerChaos 
      Caption         =   "Change"
      Height          =   285
      Left            =   9960
      TabIndex        =   41
      Top             =   5820
      Width           =   855
   End
   Begin VB.TextBox txrRunemakerChaos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9360
      TabIndex        =   40
      Text            =   "600"
      Top             =   5820
      Width           =   615
   End
   Begin VB.Timer timerSS 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   3360
   End
   Begin VB.CheckBox chkOnDangerSS 
      BackColor       =   &H00000000&
      Caption         =   "On danger screenshot"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   39
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtLowMana 
      Height          =   285
      Left            =   9360
      TabIndex        =   38
      Text            =   "100"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CheckBox chkManaFluid 
      BackColor       =   &H80000018&
      Caption         =   "Drink a mana fluid if mana is less than "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   37
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CheckBox chkmsgSound2 
      BackColor       =   &H80000018&
      Caption         =   "Play player.wav when an unfriendly player-creature pop on screen"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CheckBox chkmsgSound 
      BackColor       =   &H80000014&
      Caption         =   "Play sound PM"
      Height          =   195
      Left            =   2400
      TabIndex        =   35
      Top             =   3480
      Width           =   1515
   End
   Begin VB.CommandButton cmdStopAlarm 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stop alarms !"
      Height          =   255
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5820
      Width           =   1095
   End
   Begin VB.CheckBox chkCloseSound 
      BackColor       =   &H00000000&
      Caption         =   "Close = danger too"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   34
      Top             =   4860
      Width           =   2055
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   10800
      TabIndex        =   33
      Text            =   "fri.txt"
      ToolTipText     =   "Load - save file name"
      Top             =   2775
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000C000&
      Caption         =   "Load"
      Height          =   255
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Loads from given file"
      Top             =   2775
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      Height          =   255
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Saves to given file"
      Top             =   2775
      Width           =   615
   End
   Begin VB.OptionButton UseLeftHand 
      BackColor       =   &H00000000&
      Caption         =   "Always make runes in LEFT hand"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10980
      TabIndex        =   29
      Top             =   5640
      Width           =   3015
   End
   Begin VB.OptionButton UseRightHand 
      BackColor       =   &H00000000&
      Caption         =   "Always make runes in RIGHT hand"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10980
      TabIndex        =   28
      Top             =   5400
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton cmdDebug 
      BackColor       =   &H0080FF80&
      Caption         =   "DEBUG"
      Height          =   375
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
   End
   Begin VB.Timer TimerMaker 
      Interval        =   400
      Left            =   7320
      Top             =   3360
   End
   Begin VB.CommandButton cmdApply2 
      BackColor       =   &H80000018&
      Caption         =   "apply"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemoveFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remove sel"
      Height          =   255
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add"
      Height          =   255
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtAddFriend 
      Height          =   285
      Left            =   10920
      TabIndex        =   19
      Text            =   "friendNameHere"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ListBox lstFriends 
      Height          =   255
      Left            =   10920
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CheckBox chkLogoutOutRunes 
      BackColor       =   &H80000018&
      Caption         =   "Auto logout if out of runes or soulpoints"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CheckBox chkLogoutDangerCurrent 
      BackColor       =   &H80000018&
      Caption         =   "Auto logout if danger in current floor"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CheckBox chkLogoutDangerAny 
      BackColor       =   &H80000018&
      Caption         =   "Auto logout if danger in any floor"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   8040
      Width           =   3375
   End
   Begin VB.CheckBox chkFood 
      BackColor       =   &H80000014&
      Caption         =   "Auto eat food"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1275
   End
   Begin VB.TextBox txtSoulAction2 
      Height          =   333
      Left            =   7680
      TabIndex        =   11
      Text            =   "3"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtManaAction2 
      Height          =   333
      Left            =   8160
      TabIndex        =   9
      Text            =   "400"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtAction2 
      Height          =   333
      Left            =   6360
      TabIndex        =   7
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtManaAction1 
      Height          =   333
      Left            =   4800
      TabIndex        =   5
      Text            =   "25"
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtAction1 
      Height          =   333
      Left            =   2400
      TabIndex        =   3
      Text            =   "exura"
      Top             =   1020
      Width           =   1695
   End
   Begin VB.CheckBox chkActivate 
      BackColor       =   &H80000018&
      Caption         =   "Enable functions"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "-"
      Top             =   180
      Width           =   1815
   End
   Begin VB.CheckBox ChkDangerSound 
      BackColor       =   &H80000014&
      Caption         =   "Cavebot trapped"
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkssap 
      Caption         =   "SSA on %"
      Height          =   195
      Left            =   240
      TabIndex        =   70
      Top             =   2220
      Width           =   1095
   End
   Begin VB.CheckBox chkWaste 
      BackColor       =   &H80000014&
      Caption         =   "Mana train"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   660
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alarms"
      Height          =   1155
      Left            =   2280
      TabIndex        =   77
      Top             =   3240
      Width           =   3495
      Begin VB.CheckBox chkcd 
         Caption         =   "Cavebot danger"
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   540
         Width           =   1515
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Collectitems (use hotkey to active)"
      Height          =   255
      Left            =   7560
      TabIndex        =   80
      Top             =   3780
      Width           =   2655
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000002&
      X1              =   2280
      X2              =   5760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   6480
      X2              =   3360
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label10 
      Caption         =   "remove %"
      Height          =   255
      Left            =   3960
      TabIndex        =   74
      Top             =   2820
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "equip %"
      Height          =   255
      Left            =   2400
      TabIndex        =   73
      Top             =   2820
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   5520
      X2              =   2400
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000018&
      Caption         =   "Item ID :"
      Height          =   255
      Left            =   3480
      TabIndex        =   69
      Top             =   5400
      Width           =   855
   End
   Begin VB.Line Line18 
      BorderColor     =   &H8000000D&
      X1              =   1440
      X2              =   1440
      Y1              =   5640
      Y2              =   6480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   1440
      X2              =   4800
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line17 
      BorderColor     =   &H8000000D&
      X1              =   4800
      X2              =   4800
      Y1              =   5640
      Y2              =   6480
   End
   Begin VB.Line Line16 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   3360
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line15 
      BorderColor     =   &H8000000D&
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6480
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000D&
      X1              =   3360
      X2              =   3360
      Y1              =   6600
      Y2              =   7440
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   3360
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "Ammo ID :"
      Height          =   255
      Left            =   2400
      TabIndex        =   68
      Top             =   1920
      Width           =   855
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000D&
      X1              =   1440
      X2              =   4800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label7 
      Caption         =   "Boost Timer :"
      Height          =   255
      Left            =   6720
      TabIndex        =   57
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   2160
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000002&
      X1              =   2280
      X2              =   5760
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   2160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line d 
      BorderColor     =   &H80000002&
      X1              =   5760
      X2              =   5760
      Y1              =   600
      Y2              =   3180
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000002&
      X1              =   2280
      X2              =   2280
      Y1              =   600
      Y2              =   3180
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000002&
      X1              =   2160
      X2              =   2160
      Y1              =   600
      Y2              =   4380
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   4380
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000018&
      Caption         =   "<- lembrar"
      Height          =   255
      Left            =   9120
      TabIndex        =   45
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   " After enough mana wait up to... (ms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   43
      Top             =   6060
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Chaos (ms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   42
      Top             =   5820
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Important: you should have free space in a container called <something>BACKPACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10860
      TabIndex        =   30
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   5520
      X2              =   2400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   12360
      X2              =   10080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   10920
      X2              =   12720
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblGlobal 
      BackColor       =   &H00000000&
      Caption         =   "GLOBAL (applies for all)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   22
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Add exception:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblFriends 
      BackColor       =   &H00000000&
      Caption         =   "Don't consider following names as danger:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "soulpoints :"
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label lblMana2 
      BackColor       =   &H80000018&
      Caption         =   "mana :"
      Height          =   255
      Left            =   8640
      TabIndex        =   8
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label lblAction2 
      BackColor       =   &H80000018&
      Caption         =   "Rune Maker :"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblMana1 
      BackColor       =   &H80000014&
      Caption         =   "mana :"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000014&
      Caption         =   "Extras :"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmRunemaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Public Sub UpdateValues()

  If runemakerIDselected = 0 Then
    If RuneMakerOptions_activated_default = True Then
      chkActivate.Value = 1
    Else
      chkActivate.Value = 0
    End If
    If RuneMakerOptions_autoEat_default = True Then
      chkFood.Value = 1
    Else
      chkFood.Value = 0
    End If
    If RuneMakerOptions_ManaFluid_default = True Then
      chkManaFluid.Value = 1
    Else
      chkManaFluid.Value = 0
    End If
    If RuneMakerOptions_autoUtamo_default = True Then
      chkautoUtamo.Value = 1
    Else
      chkautoUtamo.Value = 0
    End If
    If RuneMakerOptions_autotar_default = True Then
      chkautotar.Value = 1
    Else
      chkautotar.Value = 0
    End If
    If RuneMakerOptions_autoDan_default = True Then
      chkautoDan.Value = 1
    Else
      chkautoDan.Value = 0
    End If
    If RuneMakerOptions_autodd_default = True Then
      chkautodd.Value = 1
    Else
      chkautodd.Value = 0
    End If
    If RuneMakerOptions_autoee_default = True Then
      chkautoee.Value = 1
    Else
      chkautoee.Value = 0
    End If
    If RuneMakerOptions_autoarme4_default = True Then
      chkautoarme4.Value = 1
    Else
      chkautoarme4.Value = 0
    End If
    If RuneMakerOptions_autoarme5_default = True Then
      chkautoarme5.Value = 1
    Else
      chkautoarme5.Value = 0
    End If
    If RuneMakerOptions_autoarme6_default = True Then
      chkautoarme6.Value = 1
    Else
      chkautoarme6.Value = 0
    End If
    If RuneMakerOptions_autora_default = True Then
      chkautora.Value = 1
    Else
      chkautora.Value = 0
    End If
    If RuneMakerOptions_autoda_default = True Then
      chkautoda.Value = 1
    Else
      chkautoda.Value = 0
    End If
    If RuneMakerOptions_autoxray_default = True Then
      chkautoxray.Value = 1
    Else
      chkautoxray.Value = 0
    End If
    If RuneMakerOptions_autodk_default = True Then
      chkautodk.Value = 1
    Else
      chkautodk.Value = 0
    End If
    If RuneMakerOptions_autogHur_default = True Then
      chkautogHur.Value = 1
    Else
      chkautogHur.Value = 0
    End If
    If RuneMakerOptions_autoHur_default = True Then
      chkautoHur.Value = 1
    Else
      chkautoHur.Value = 0
    End If
    If RuneMakerOptions_autoPM2_default = True Then
      chkautoPM2.Value = 1
    Else
      chkautoPM2.Value = 0
    End If
    If RuneMakerOptions_autoLogoutAnyFloor_default = True Then
      chkLogoutDangerAny.Value = 1
    Else
      chkLogoutDangerAny.Value = 0
    End If
    If RuneMakerOptions_autoLogoutCurrentFloor_default = True Then
      chkLogoutDangerCurrent.Value = 1
    Else
      chkLogoutDangerCurrent.Value = 0
    End If
    If RuneMakerOptions_autoLogoutOutOfRunes_default = True Then
      chkLogoutOutRunes.Value = 1
    Else
      chkLogoutOutRunes.Value = 0
    End If
    If RuneMakerOptions_autoWaste_default = True Then
      chkWaste.Value = 1
    Else
      chkWaste.Value = 0
    End If
    If RuneMakerOptions_autossap_default = True Then
      chkssap.Value = 1
    Else
      chkssap.Value = 0
    End If
    If RuneMakerOptions_autoerg_default = True Then
      chkerg.Value = 1
    Else
      chkerg.Value = 0
    End If
    If RuneMakerOptions_msgSound_default = True Then
      chkmsgSound.Value = 1
    Else
      chkmsgSound.Value = 0
    End If
    If RuneMakerOptions_msgSound2_default = True Then
      chkmsgSound2.Value = 1
    Else
      chkmsgSound2.Value = 0
    End If
    txtAction1.Text = RuneMakerOptions_firstActionText_default
    Text1.Text = CStr(RuneMakerOptions_thirdActionText_default)
    txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
    txtbeep.Text = RuneMakerOptions_beeploot_default
    Text2.Text = CStr(RuneMakerOptions_text2_default)
    Text3.Text = CStr(RuneMakerOptions_text3_default)
    txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
    txtAction2.Text = RuneMakerOptions_secondActionText_default
    txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
    txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
  Else
  'aq tentativa ativar refil e collect
    'If TrainerOptions(runemakerIDselected).enabled = True Then
    '  chkEnableTrainer.Value = 1
    'Else
    '  chkEnableTrainer.Value = 0
    'End If
    'If TrainerOptions(runemakerIDselected).PlayerSlots(10).cheked = True Then
    '  chkSlotRefill(10).Value = 1
    'Else
    '  chkSlotRefill(10).Value = 0
    'End If
    If RuneMakerOptions(runemakerIDselected).activated = True Then
      chkActivate.Value = 1
    Else
      chkActivate.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoEat = True Then
      chkFood.Value = 1
    Else
      chkFood.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).ManaFluid = True Then
      chkManaFluid.Value = 1
    Else
      chkManaFluid.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoUtamo = True Then
      chkautoUtamo.Value = 1
    Else
      chkautoUtamo.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autotar = True Then
      chkautotar.Value = 1
    Else
      chkautotar.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autopmax = True Then
      chkautopmax.Value = 1
    Else
      chkautopmax.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoAp = True Then
      chkautoAp.Value = 1
    Else
      chkautoAp.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autossa = True Then
      chkautossa.Value = 1
    Else
      chkautossa.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoSdt = True Then
      chkautoSdt.Value = 1
    Else
      chkautoSdt.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoDan = True Then
      chkautoDan.Value = 1
    Else
      chkautoDan.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autodd = True Then
      chkautodd.Value = 1
    Else
      chkautodd.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoee = True Then
      chkautoee.Value = 1
    Else
      chkautoee.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoarme4 = True Then
      chkautoarme4.Value = 1
    Else
      chkautoarme4.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoarme5 = True Then
      chkautoarme5.Value = 1
    Else
      chkautoarme5.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoarme6 = True Then
      chkautoarme6.Value = 1
    Else
      chkautoarme6.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autora = True Then
      chkautora.Value = 1
    Else
      chkautora.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoda = True Then
      chkautoda.Value = 1
    Else
      chkautoda.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoxray = True Then
      chkautoxray.Value = 1
    Else
      chkautoxray.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autodk = True Then
      chkautodk.Value = 1
    Else
      chkautodk.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autogHur = True Then
      chkautogHur.Value = 1
    Else
      chkautogHur.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoHur = True Then
      chkautoHur.Value = 1
    Else
      chkautoHur.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoPM2 = True Then
      chkautoPM2.Value = 1
    Else
      chkautoPM2.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = True Then
      chkLogoutDangerAny.Value = 1
    Else
      chkLogoutDangerAny.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = True Then
      chkLogoutDangerCurrent.Value = 1
    Else
      chkLogoutDangerCurrent.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = True Then
      chkLogoutOutRunes.Value = 1
    Else
      chkLogoutOutRunes.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoWaste = True Then
      chkWaste.Value = 1
    Else
      chkWaste.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autossap = True Then
      chkssap.Value = 1
    Else
      chkssap.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoerg = True Then
      chkerg.Value = 1
    Else
      chkerg.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).msgSound = True Then
      chkmsgSound.Value = 1
    Else
      chkmsgSound.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).msgSound2 = True Then
      chkmsgSound2.Value = 1
    Else
      chkmsgSound2.Value = 0
    End If
    txtAction1.Text = RuneMakerOptions(runemakerIDselected).firstActionText
    Text1.Text = CStr(RuneMakerOptions(runemakerIDselected).thirdActionText)
    txtManaAction1.Text = CStr(RuneMakerOptions(runemakerIDselected).firstActionMana)
    txtbeep.Text = RuneMakerOptions(runemakerIDselected).beeploot
    Text2.Text = CStr(RuneMakerOptions(runemakerIDselected).Text2)
    txtLowMana.Text = CStr(RuneMakerOptions(runemakerIDselected).LowMana)
    txtAction2.Text = RuneMakerOptions(runemakerIDselected).secondActionText
    txtManaAction2.Text = CStr(RuneMakerOptions(runemakerIDselected).secondActionMana)
    txtSoulAction2.Text = CStr(RuneMakerOptions(runemakerIDselected).secondActionSoulpoints)
  End If
End Sub

Public Sub SetChk(typeChk As String, v As Integer)
  Select Case typeChk
  Case "chkActivate"
    lock_chkActivate = True
    chkActivate.Value = v
    lock_chkActivate = False
    
  Case "chkFood"
    lock_chkFood = True
    chkFood.Value = v
    lock_chkFood = False
    
  Case "chkManaFluid"
    lock_chkManaFluid = True
    chkManaFluid.Value = v
    lock_chkManaFluid = False
    
  Case "chkautoUtamo"
    lock_chkautoUtamo = True
    chkautoUtamo.Value = v
    lock_chkautoUtamo = False
    
  Case "chkautotar"
    lock_chkautotar = True
    chkautotar.Value = v
    lock_chkautotar = False
    
  Case "chkautoAp"
    lock_chkautoAp = True
    chkautoAp.Value = v
    lock_chkautoAp = False
    
    Case "chkautossa"
    lock_chkautossa = True
    chkautossa.Value = v
    lock_chkautossa = False
    
  Case "chkautopmax"
    lock_chkautopmax = True
    chkautopmax.Value = v
    lock_chkautopmax = False
    
  Case "chkautoSdt"
    lock_chkautoSdt = True
    chkautoSdt.Value = v
    lock_chkautoSdt = False
    
  Case "chkautoDan"
    lock_chkautoDan = True
    chkautoDan.Value = v
    lock_chkautoDan = False
    
  Case "chkautodd"
    lock_chkautodd = True
    chkautodd.Value = v
    lock_chkautodd = False
    
    Case "chkautoee"
    lock_chkautoee = True
    chkautoee.Value = v
    lock_chkautoee = False
    
    Case "chkautoarme4"
    lock_chkautoarme4 = True
    chkautoarme4.Value = v
    lock_chkautoarme4 = False
    
    Case "chkautoarme5"
    lock_chkautoarme5 = True
    chkautoarme5.Value = v
    lock_chkautoarme5 = False
    
    Case "chkautoarme6"
    lock_chkautoarme6 = True
    chkautoarme6.Value = v
    lock_chkautoarme6 = False
    
    Case "chkautora"
    lock_chkautora = True
    chkautora.Value = v
    lock_chkautora = False
    
    Case "chkautoda"
    lock_chkautoda = True
    chkautoda.Value = v
    lock_chkautoda = False
    
  Case "chkautoxray"
    lock_chkautoxray = True
    chkautoxray.Value = v
    lock_chkautoxray = False
    
  Case "chkautodk"
    lock_chkautodk = True
    chkautodk.Value = v
    lock_chkautodk = False

  Case "chkautogHur"
    lock_chkautogHur = True
    chkautogHur.Value = v
    lock_chkautogHur = False
    
  Case "chkautoHur"
    lock_chkautoHur = True
    chkautoHur.Value = v
    lock_chkautoHur = False
    
  Case "chkautoPM2"
    lock_chkautoPM2 = True
    chkautoPM2.Value = v
    lock_chkautoPM2 = False
    
  Case "chkautoaim"
    lock_chkautoaim = True
    frmStealth.chkautoaim.Value = v
    lock_chkautoaim = False

  Case "chkautoUE"
    lock_chkautoUE = True
    frmStealth.chkautoUE.Value = v
    lock_chkautoUE = False

  Case "chklocktrigger"
    lock_chklocktrigger = True
    frmStealth.chklocktrigger.Value = v
    lock_chklocktrigger = False
    
  Case "chkLogoutDangerAny"
    lock_chkLogoutDangerAny = True
    chkLogoutDangerAny.Value = v
    lock_chkLogoutDangerAny = False
    
  Case "chkLogoutDangerCurrent"
    lock_chkLogoutDangerCurrent = True
    chkLogoutDangerCurrent.Value = v
    lock_chkLogoutDangerCurrent = False
    
  Case "chkLogoutOutRunes"
    lock_chkLogoutOutRunes = True
    chkLogoutOutRunes.Value = v
    lock_chkLogoutOutRunes = False
    
  Case "chkWaste"
    lock_chkWaste = True
    chkWaste.Value = v
    lock_chkWaste = False
    
  Case "chkssap"
    lock_chkssap = True
    chkssap.Value = v
    lock_chkssap = False
    
  Case "chkerg"
    lock_chkerg = True
    chkerg.Value = v
    lock_chkerg = False
    
  Case "chkmsgSound"
    lock_chkmsgSound = True
    chkmsgSound.Value = v
    lock_chkmsgSound = False
  
  Case "chkmsgSound2"
    lock_chkmsgSound2 = True
    chkmsgSound2.Value = v
    lock_chkmsgSound2 = False
    
  Case "chkAutoVita2"
    lock_chkAutoVita2 = True
    frmHardcoreCheats.chkAutoVita2.Value = v
    lock_chkAutoVita2 = False

  Case "chkAutoVita"
    lock_chkAutoVita = True
    frmHardcoreCheats.chkAutoVita.Value = v
    lock_chkAutoVita = False

  Case "chkAutoVita4"
    lock_chkAutoVita4 = True
    frmHardcoreCheats.chkAutoVita4.Value = v
    lock_chkAutoVita4 = False

  Case "chkAutoVita3"
    lock_chkAutoVita3 = True
    frmHardcoreCheats.chkAutoVita3.Value = v
    lock_chkAutoVita3 = False
    
  Case "chkarme"
    lock_chkarme = True
    frmHardcoreCheats.Check1.Value = v
    lock_chkarme = False

    Case "chkarme2"
    lock_chkarme2 = True
    frmHardcoreCheats.Check2.Value = v
    lock_chkarme2 = False
    
    Case "chkarme3"
    lock_chkarme3 = True
    frmHardcoreCheats.Check3.Value = v
    lock_chkarme3 = False
    
  End Select
End Sub

Public Sub DisableAll(id As Integer)
  If id = CInt(runemakerIDselected) Then
    SetChk "chkFood", 0
    SetChk "chkManaFluid", 0
    SetChk "chkautoUtamo", 0
    SetChk "chkautotar", 0
    SetChk "chkautoAp", 0
    SetChk "chkautossa", 0
    SetChk "chkautopmax", 0
    SetChk "chkautoSdt", 0
    SetChk "chkautoDan", 0
    SetChk "chkautodd", 0
    SetChk "chkautoee", 0
    SetChk "chkautoarme4", 0
    SetChk "chkautoarme5", 0
    SetChk "chkautoarme6", 0
    SetChk "chkautora", 0
    SetChk "chkautoda", 0
    SetChk "chkautoxray", 0
    SetChk "chkautodk", 0
    SetChk "chkautogHur", 0
    SetChk "chkautoHur", 0
    SetChk "chkautoPM2", 0
    SetChk "chkautoaim", 0
    SetChk "chkautoUE", 0
    SetChk "chklocktrigger", 0
    SetChk "chkLogoutDangerAny", 0
    SetChk "chkLogoutDangerCurrent", 0
    SetChk "chkLogoutOutRunes", 0
    SetChk "chkWaste", 0
    SetChk "chkssap", 0
    SetChk "chkerg", 0
    SetChk "chkmsgSound", 0
    SetChk "chkmsgSound2", 0
  End If
  
  If id = CInt(HardcoreCheatsIDselected) Then
    SetChk "chkAutoVita2", 0
    SetChk "chkAutoVita", 0
    SetChk "chkAutoVita4", 0
    SetChk "chkAutoVita3", 0
    SetChk "chkCheck1", 0
  End If
  
  'HardcoreCheatsOptions(id).txtExuraVita2 = True
  HardcoreCheatsOptions(id).arme = False
  HardcoreCheatsOptions(id).arme2 = False
  HardcoreCheatsOptions(id).arme3 = False
  HardcoreCheatsOptions(id).sphi = False
  HardcoreCheatsOptions(id).splo = False
  HardcoreCheatsOptions(id).pmh = False
  HardcoreCheatsOptions(id).pth = False

  RuneMakerOptions(id).activated = True
  RuneMakerOptions(id).autoEat = False
  RuneMakerOptions(id).ManaFluid = False
  RuneMakerOptions(id).autoUtamo = False
  RuneMakerOptions(id).autotar = False
  RuneMakerOptions(id).autoAp = False
  RuneMakerOptions(id).autossa = False
  RuneMakerOptions(id).autopmax = False
  RuneMakerOptions(id).autoSdt = False
  RuneMakerOptions(id).autoDan = False
  RuneMakerOptions(id).autodd = False
  RuneMakerOptions(id).autoee = False
  RuneMakerOptions(id).autoarme4 = False
  RuneMakerOptions(id).autoarme5 = False
  RuneMakerOptions(id).autoarme6 = False
  RuneMakerOptions(id).autora = False
  RuneMakerOptions(id).autoda = False
  RuneMakerOptions(id).autoxray = False
  RuneMakerOptions(id).autodk = False
  RuneMakerOptions(id).autogHur = False
  RuneMakerOptions(id).autoHur = False
  RuneMakerOptions(id).autoPM2 = False
  RuneMakerOptions(id).autoaim = False
  RuneMakerOptions(id).autoUE = False
  RuneMakerOptions(id).locktrigger = False
  RuneMakerOptions(id).autoLogoutAnyFloor = False
  RuneMakerOptions(id).autoLogoutCurrentFloor = False
  RuneMakerOptions(id).autoLogoutOutOfRunes = False
  RuneMakerOptions(id).autoWaste = False
  RuneMakerOptions(id).autossap = False
  RuneMakerOptions(id).autoerg = False
  RuneMakerOptions(id).msgSound = False
  RuneMakerOptions(id).msgSound2 = False
End Sub










Private Sub Check1_Click()
Check1.Value = 2
End Sub









Private Sub chkautoee_Click()
'If lock_chkautoee = False Then
'If trainerIDselected > 0 Then
'  If chkautoee.Value = 1 Then
'    RuneMakerOptions(trainerIDselected).autoee = True
'  Else
'    RuneMakerOptions(trainerIDselected).autoee = False
'  End If
'End If
'End If
End Sub

Private Sub chkautodd_Click()
Dim idConnection As Integer
    
If lock_chkautodd = False Then
If runemakerIDselected > 0 Then
  If chkautodd.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autodd = True
  Else
    RuneMakerOptions(runemakerIDselected).autodd = False
  End If
End If
End If

End Sub

Private Sub chkautodk_Click()
Dim idConnection As Integer
    
If lock_chkautodk = False Then
If runemakerIDselected > 0 Then
  If chkautodk.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autodk = True
  Else
    RuneMakerOptions(runemakerIDselected).autodk = False
  End If
End If
End If

End Sub





Private Sub chkautoPM2_Click()
Dim idConnection As Integer

If lock_chkautoPM2 = False Then
If runemakerIDselected > 0 Then
  If chkautoPM2.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoPM2 = True
  Else
    RuneMakerOptions(runemakerIDselected).autoPM2 = False
  End If
End If
End If
End Sub

Private Sub chkautoxray_Click()
Dim idConnection As Integer
   
If lock_chkautoxray = False Then
If runemakerIDselected > 0 Then
  If chkautoxray.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoxray = True
  Else
    RuneMakerOptions(runemakerIDselected).autoxray = False
  End If
End If
End If

End Sub







Private Sub chkerg_Click()
If lock_chkerg = False Then
If runemakerIDselected > 0 Then
  If chkerg.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoerg = True
  Else
    RuneMakerOptions(runemakerIDselected).autoerg = False
  End If
End If
End If
End Sub

Private Sub chkActivate_Click()
Dim tileID As Long
Dim aRes As Long
#If FinalMode Then
On Error GoTo goterr
#End If



If lock_chkActivate = False Then
If runemakerIDselected > 0 Then
  If chkActivate.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).activated = True
    If TibiaVersionLong >= 872 Then
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t3
      End If
      ' aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started.")
      DoEvents
    Else
    If UseRightHand.Value = True Then
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on right hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2))
        DoEvents
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t3
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on right hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2) & " (with amount byte " & GoodHex(savedItem(runemakerIDselected).t3) & " )")
        DoEvents
      End If
    Else
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on left hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2))
        DoEvents
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t3
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on left hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2) & " (with amount byte  " & GoodHex(savedItem(runemakerIDselected).t3) & " )")
        DoEvents
      End If
    End If
    End If
  Else
  'aqui tentativa de habilitar essa porra
    RuneMakerOptions(runemakerIDselected).activated = True
  End If
End If
End If
Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Warning: connection fail during the runemaker activation - ignoring"
End Sub





Private Sub chkautoAp_Click()
If lock_chkautoAp = False Then
If runemakerIDselected > 0 Then
  If chkautoAp.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoAp = True
  Else
    RuneMakerOptions(runemakerIDselected).autoAp = False
  End If
End If
End If
End Sub

Private Sub chkautoDan_Click()
Dim idConnection As Integer
    
If lock_chkautoDan = False Then
If runemakerIDselected > 0 Then
  If chkautoDan.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoDan = True
  Else
    RuneMakerOptions(runemakerIDselected).autoDan = False
  End If
End If
End If

End Sub

Private Sub chkautogHur_Click()
Dim idConnection As Integer

If lock_chkautogHur = False Then
If runemakerIDselected > 0 Then
  If chkautogHur.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autogHur = True
  Else
    RuneMakerOptions(runemakerIDselected).autogHur = False
  End If
End If
End If

End Sub

Private Sub chkautoHur_Click()
Dim idConnection As Integer
Dim aRes As Long

If lock_chkautoHur = False Then
If runemakerIDselected > 0 Then
  If chkautoHur.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoHur = True
  Else
    RuneMakerOptions(runemakerIDselected).autoHur = False
  End If
End If
End If

End Sub

Private Sub chkautopmax_Click()
If lock_chkautopmax = False Then
If runemakerIDselected > 0 Then
  If chkautopmax.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autopmax = True
  Else
    RuneMakerOptions(runemakerIDselected).autopmax = False
  End If
End If
End If
End Sub

Private Sub chkautoSdt_Click()
If lock_chkautoSdt = False Then
If runemakerIDselected > 0 Then
  If chkautoSdt.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoSdt = True
  Else
    RuneMakerOptions(runemakerIDselected).autoSdt = False
  End If
End If
End If
End Sub



Private Sub chkautossa_Click()
If lock_chkautossa = False Then
If runemakerIDselected > 0 Then
  If chkautossa.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autossa = True
  Else
    RuneMakerOptions(runemakerIDselected).autossa = False
  End If
End If
End If
End Sub

Private Sub chkautoUtamo_Click()
Dim idConnection As Integer
    
If lock_chkautoUtamo = False Then
If runemakerIDselected > 0 Then
  If chkautoUtamo.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoUtamo = True
  Else
    RuneMakerOptions(runemakerIDselected).autoUtamo = False
  End If
End If
End If

End Sub







Private Sub chkEnableTrainer_Click()

  If (runemakerIDselected > 0) Then
    TrainerOptions(runemakerIDselected).enabled = chkEnableTrainer.Value
  End If
End Sub

Private Sub chkFood_Click()
If lock_chkFood = False Then
If runemakerIDselected > 0 Then
  If chkFood.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoEat = True
  Else
    RuneMakerOptions(runemakerIDselected).autoEat = False
  End If
End If
End If
End Sub

Private Sub chkLogoutDangerAny_Click()
If lock_chkLogoutDangerAny = False Then
If runemakerIDselected > 0 Then
  If chkLogoutDangerAny.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = True
    SetChk "chkLogoutDangerCurrent", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = False
  End If
End If
End If
End Sub

Private Sub chkLogoutDangerCurrent_Click()
If lock_chkLogoutDangerCurrent = False Then
If runemakerIDselected > 0 Then
  If chkLogoutDangerCurrent.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = True
    SetChk "chkLogoutDangerAny", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = False
  End If
End If
End If
End Sub

Private Sub chkLogoutOutRunes_Click()
If lock_chkLogoutOutRunes = False Then
If runemakerIDselected > 0 Then
  If chkLogoutOutRunes.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = True
    'SetChk "chkWaste", 0
    'RuneMakerOptions(runemakerIDselected).autoWaste = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = False
  End If
End If
End If
End Sub

Private Sub chkManaFluid_Click()
If lock_chkManaFluid = False Then
If runemakerIDselected > 0 Then
  If chkManaFluid.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).ManaFluid = True
  Else
    RuneMakerOptions(runemakerIDselected).ManaFluid = False
    RemoveSpamOrder CInt(runemakerIDselected), 4 'remove auto mana
  End If
End If
End If
End Sub

Private Sub chkmsgSound2_Click()
If lock_chkmsgSound2 = False Then
If runemakerIDselected > 0 Then
  If chkmsgSound2.Value = 1 Then
    DangerPlayer(runemakerIDselected) = False
    RuneMakerOptions(runemakerIDselected).msgSound2 = True
  Else
    DangerPlayer(runemakerIDselected) = False
    RuneMakerOptions(runemakerIDselected).msgSound2 = False
  End If
End If
End If
End Sub





Private Sub chkReveal_Click()
chkReveal.Value = 1
chkReveal.enabled = False
End Sub



Private Sub chkautotar_Click()
Dim idConnection As Integer
Dim aRes As Long

If lock_chkautotar = False Then
If runemakerIDselected > 0 Then
  If chkautotar.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autotar = True
  Else
    RuneMakerOptions(runemakerIDselected).autotar = False
  End If
End If
End If
    
    'For idConnection = 1 To MAXCLIENTS
    'If (GameConnected(idConnection) = True) Then
    'If chktar.Value = 1 Then
    'RuneMakerOptions(idConnection).autotar = True
    'aRes = SendCustomSystemMessageToClient(idConnection, "Hold Target ON", &HB)
    'Timertar.enabled = True
    'Else
    'RuneMakerOptions(idConnection).autotar = False
    'aRes = SendCustomSystemMessageToClient(idConnection, "Hold Target OFF", &HB)
    'Timertar.enabled = False
    'End If
    '  End If
    'Next idConnection

End Sub



Private Sub chkSlotRefill_Click(index As Integer)
  Dim thenewvalue As Long
  thenewvalue = chkSlotRefill(index).Value
  If ((runemakerIDselected > 0) And (index > 0)) Then
    TrainerOptions(runemakerIDselected).PlayerSlots(index).cheked = thenewvalue
  End If
End Sub

Private Sub chkssap_Click()
Dim idConnection As Integer
    
If lock_chkssap = False Then
If runemakerIDselected > 0 Then
  If chkssap.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autossap = True
  Else
    RuneMakerOptions(runemakerIDselected).autossap = False
  End If
End If
End If

End Sub

Private Sub chkWaste_Click()
If lock_chkWaste = False Then
If runemakerIDselected > 0 Then
  If chkWaste.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoWaste = True
    SetChk "chkLogoutOutRunes", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = False
  Else
    RuneMakerOptions(runemakerIDselected).autoWaste = False
  End If
End If
End If
End Sub


Private Sub chkmsgSound_Click()
If lock_chkmsgSound = False Then
If runemakerIDselected > 0 Then
  If chkmsgSound.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).msgSound = True
  Else
    RuneMakerOptions(runemakerIDselected).msgSound = False
  End If
End If
End If
End Sub



Private Sub cmbCharacter_Click()
 runemakerIDselected = cmbCharacter.ListIndex
   If runemakerIDselected = 0 Then
      runemakerIDselected = 1
    '  Exit Sub
  End If
  If runemakerIDselected > 0 Then
      UpdateValues
      'frmTrainer.UpdateValues
  'End If
  'aq tentativa
  
    If TrainerOptions(runemakerIDselected).enabled = 1 Then
    chkEnableTrainer.Value = 1
    Else
    chkEnableTrainer.Value = 0
    End If
    If TrainerOptions(runemakerIDselected).PlayerSlots(10).cheked = 1 Then
      chkSlotRefill(10).Value = 1
    Else
      chkSlotRefill(10).Value = 0
    End If
    
    End If

End Sub
Public Function IsFriend(strName As String) As Boolean
  'strname should come in lcase
  Dim i As Long
  Dim totI As Long
  Dim foundI As Long
  totI = lstFriends.ListCount - 1
  foundI = -1
  For i = 0 To totI
    If lstFriends.List(i) = strName Then
      foundI = i
    End If
  Next i
  If foundI = -1 Then
    IsFriend = False
  Else
    IsFriend = True
  End If
End Function

Private Sub cmdAddFriend_Click()
  If IsFriend(LCase(txtAddFriend.Text)) = False Then
    lstFriends.AddItem LCase(txtAddFriend.Text)
  End If
End Sub

Private Sub cmdApply_Click()

    UpdateValues
 
End Sub

Private Sub cmdDebug_Click()
  Dim aRes As Long
  Dim i As Long
  If runemakerIDselected > 0 Then
    publicDebugMode = Not publicDebugMode
    If publicDebugMode = True Then
      aRes = GiveGMmessage(CInt(runemakerIDselected), "DEBUG MODE ENABLED", "Blackd")
      For i = 1 To MAXCLIENTS
        If GameConnected(i) = True Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Process ID of client #" & CStr(i) & " (" & CharacterName(i) & ") =" & CStr(ProcessID(i))
        End If
      Next i
      DoEvents
    Else
      aRes = GiveGMmessage(CInt(runemakerIDselected), "DEBUG MODE DISABLED", "Blackd")
      DoEvents
    End If
  End If
End Sub

Private Sub cmdLoad_Click()
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim filename As String
  Set fso = New scripting.FileSystemObject
    lstFriends.Clear
    filename = App.path & "\" & txtFile.Text
    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        If strLine <> "" Then
        If IsFriend(LCase(strLine)) = False Then
          lstFriends.AddItem LCase(strLine)
        End If
        End If
      Wend
      Close #fn
    End If
End Sub

Private Sub cmdRemoveFriend_Click()
  If lstFriends.ListIndex > -1 Then
    lstFriends.RemoveItem (lstFriends.ListIndex)
  End If
End Sub

Private Sub cmdSave_Click()
  Dim fn As Integer
  Dim limI As Long
  Dim i As Long
    limI = lstFriends.ListCount - 1
    fn = FreeFile
    Open App.path & "\" & txtFile.Text For Output As #fn
    For i = 0 To limI
      Print #fn, lstFriends.List(i)
    Next i
    Close #fn
End Sub

Private Sub cmdSaveRunemakerChaos_Click()
    On Error GoTo goterr
    Dim lngCast As Long
    Dim lngCast2 As Long
    lngCast = CLng(frmRunemaker.txrRunemakerChaos.Text)
    lngCast2 = CLng(frmRunemaker.txrRunemakerChaos2.Text)
    If (lngCast >= 20) And (lngCast2 >= 0) Then
        RunemakerChaos = lngCast
        RunemakerChaos2 = lngCast2
        Me.txrRunemakerChaos.Text = CStr(RunemakerChaos)
        Me.txrRunemakerChaos2.Text = CStr(RunemakerChaos2)
        frmRunemaker.Caption = "Runemaker - chaos updated"
    Else
        GoTo goterr
    End If
    Exit Sub
goterr:
    frmRunemaker.Caption = "Runemaker - invalid chaos values"
End Sub

Private Sub cmdStopAlarm_Click()
  Dim mcid As Integer
  For mcid = 1 To MAXCLIENTS
    DangerPK(mcid) = False
    DangerGM(mcid) = False
    LogoutTimeGM(mcid) = 0
    moveRetry(mcid) = 0
    RemoveSpamOrder mcid, 1
    UHRetryCount(mcid) = 0
  Next mcid
  ChangePlayTheDangerSound False
End Sub



Private Sub Combo1_Click()
Dim index As Integer
'Dim Index As Integer
'Dim i As Integer
'ammo.ListIndex = ammo.NewIndex
If Combo1.ListIndex = 0 Then
'Text1.Text = "77 0D" 'arrow
txtSlotRefill(10).Text = "77 0D"
ElseIf Combo1.ListIndex = 1 Then
'Text1.Text = "79 0D" 'burst
txtSlotRefill(10).Text = "79 0D"
ElseIf Combo1.ListIndex = 2 Then
'Text1.Text = "C4 1C" 'sniper
txtSlotRefill(10).Text = "C4 1C"
ElseIf Combo1.ListIndex = 3 Then
'Text1.Text = "AB 37" 'tarsal
txtSlotRefill(10).Text = "AB 37"
ElseIf Combo1.ListIndex = 4 Then
'Text1.Text = "C5 1C" 'onyx
txtSlotRefill(10).Text = "C5 1C"
ElseIf Combo1.ListIndex = 5 Then
'Text1.Text = "0F 3F" 'envenon
txtSlotRefill(10).Text = "0F 3F"
ElseIf Combo1.ListIndex = 6 Then
'Text1.Text = "B1 3D" 'crystal
txtSlotRefill(10).Text = "B1 3D"
ElseIf Combo1.ListIndex = 7 Then
'Text1.Text = "76 0D" ' bolt
txtSlotRefill(10).Text = "76 0D"
ElseIf Combo1.ListIndex = 8 Then
'Text1.Text = "C3 1C" 'piercing
txtSlotRefill(10).Text = "C3 1C"
ElseIf Combo1.ListIndex = 9 Then
'Text1.Text = "AC 37" 'vortex
txtSlotRefill(10).Text = "AC 37"
ElseIf Combo1.ListIndex = 10 Then
'Text1.Text = "0E 3F" 'drill
txtSlotRefill(10).Text = "0E 3F"
ElseIf Combo1.ListIndex = 11 Then
'Text1.Text = "7A 0D" 'pbolt
txtSlotRefill(10).Text = "7A 0D"
ElseIf Combo1.ListIndex = 12 Then
'Text1.Text = "0D 3F" 'prismatic
txtSlotRefill(10).Text = "0D 3F"
ElseIf Combo1.ListIndex = 13 Then
'Text1.Text = "80 19" 'infernal
txtSlotRefill(10).Text = "80 19"
End If

End Sub

Private Sub Command1_Click()
  frmTrainer.WindowState = vbNormal
  frmTrainer.Show
  frmTrainer.SetFocus
End Sub

Private Sub Command2_Click()
  frmIdlist.WindowState = vbNormal
  frmIdlist.Show
  frmIdlist.SetFocus
  SetWindowPos frmIdlist.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Load()
Combo1.AddItem "Arrow"
Combo1.AddItem "Burst Arrow"
Combo1.AddItem "Sniper Arrow"
Combo1.AddItem "Tarsal Arrow"
Combo1.AddItem "Onyx Arrow"
Combo1.AddItem "Envenomed Arrow"
Combo1.AddItem "Crystalline Arrow"
Combo1.AddItem "Bolt"
Combo1.AddItem "Piercing Bolt"
Combo1.AddItem "Power Bolt"
Combo1.AddItem "Vortex Bolt"
Combo1.AddItem "Drill Bolt"
Combo1.AddItem "Prismatic Bolt"
Combo1.AddItem "Infernal Bolt"
LoadRuneChars
Check1.Value = 2
Check2.Value = 2
chkActivate.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub
Public Sub LoadRuneChars()
  Dim i As Long
  Dim firstC As Long
  If TibiaVersionLong >= 872 Then
    UseRightHand.Visible = False
    UseLeftHand.Visible = False
    'fraNoHands.Visible = True
    Label2.Caption = "IMPORTANT: " & currentAppName & " will only count runes displayed in opened backpacks!"
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
  runemakerIDselected = firstC
  UpdateValues
End Sub








Private Sub Text1_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(Text1.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).thirdActionText = lonN
  Else
    Text1.Text = CStr(RuneMakerOptions_thirdActionText_default)
    RuneMakerOptions(runemakerIDselected).thirdActionText = RuneMakerOptions_thirdActionText_default
  End If
  End If
  Exit Sub
gotError:
  Text1.Text = CStr(RuneMakerOptions_thirdActionText_default)
  RuneMakerOptions(runemakerIDselected).thirdActionText = RuneMakerOptions_thirdActionText_default
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(Text2.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).Text2 = lonN
  Else
    Text2.Text = CStr(RuneMakerOptions_text2_default)
    RuneMakerOptions(runemakerIDselected).Text2 = RuneMakerOptions_text2_default
  End If
  End If
  Exit Sub
gotError:
  Text2.Text = CStr(RuneMakerOptions_text2_default)
  RuneMakerOptions(runemakerIDselected).Text2 = RuneMakerOptions_text2_default
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(Text3.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).Text3 = lonN
  Else
    Text3.Text = CStr(RuneMakerOptions_text3_default)
    RuneMakerOptions(runemakerIDselected).Text3 = RuneMakerOptions_text3_default
  End If
  End If
  Exit Sub
gotError:
  Text3.Text = CStr(RuneMakerOptions_text3_default)
  RuneMakerOptions(runemakerIDselected).Text3 = RuneMakerOptions_text3_default
End Sub




Private Sub TimerAp_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoAp = True) Then
                aRes = ExecuteInTibia("exiva drop D7 0B 01", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub TimerDan2_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoDan = True) Then
                aRes = ExecuteInTibia("exiva dance", idConnection, True)
            End If
      End If
    Next idConnection
End Sub





Private Sub Timererg_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim percent As Long
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       percent = 100 * ((myHP(idConnection) / myMaxHP(idConnection)))
      If (RuneMakerOptions(idConnection).autoerg = True) And (percent < RuneMakerOptions(idConnection).Text2) Then
                aRes = ExecuteInTibia("exiva #EB 0B 09", idConnection, True)
            End If
            If (RuneMakerOptions(idConnection).autoerg = True) And (percent > RuneMakerOptions(idConnection).Text3) Then
                aRes = ExecuteInTibia("Exiva > 78 FF FF 09 00 00 $hex-equiped-item:09$ 00 FF FF 03 00 00 01", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub TimergHur2_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autogHur = True) And (GetStatusBit(idConnection, 2) = 0) Then
                aRes = ExecuteInTibia("utani gran hur", idConnection, True)
            End If
      End If
    Next idConnection
End Sub



Private Sub TimerHur2_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoHur = True) And (GetStatusBit(idConnection, 2) = 0) Then
                aRes = ExecuteInTibia("utani hur", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub TimerMaker_Timer()
  Dim utm As Long
  Dim resS As TypeSearchItemResult2
  Dim idConnection As Integer
  Dim aRes As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim cond1 As Boolean
  Dim cond2 As Boolean
  Dim cond3 As Boolean
  Dim cond4 As Boolean
  Dim cond5 As Boolean
  Dim tmpcond As Boolean
  Dim playerS As String
  Dim gtc As Long
  Dim eatDone() As Boolean
  Dim i As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  'add chaos to the timer
  'TimerMaker.Interval = 400 + randomNumberBetween(0, RunemakerChaos)
      
  ' EAT FOOD EACH 60 TURNS (1 turn = 400ms)
  ' when eating food > do nothing else in this mc
  ReDim eatDone(1 To MAXCLIENTS)
  For i = 1 To MAXCLIENTS
    eatDone(i) = False
  Next i
  For idConnection = 1 To MAXCLIENTS
    If (runeTurn(idConnection) > 59) Then
        If ((GameConnected(idConnection) = True) And (sentWelcome(idConnection) = True) And (GotPacketWarning(idConnection) = False)) Then
          If makingRune(idConnection) = False Then
              If RuneMakerOptions(idConnection).autoEat = True Then
               ' We are allowed to eat.
               ' Lets search food...
                resS = SearchFood(idConnection)
                If (resS.foundcount > 0) Then
                  ' Food found, eat it now
                  aRes = EatFood(idConnection, resS.b1, resS.b2, resS.bpID, resS.slotID)
                  DoEvents
                End If
              End If
          End If
        End If
        runeTurn(idConnection) = randomNumberBetween(0, 29)
        eatDone(idConnection) = True
    Else
      runeTurn(idConnection) = runeTurn(idConnection) + 1
    End If
  Next idConnection
  
  gtc = GetTickCount()
  For idConnection = 1 To MAXCLIENTS
    If (eatDone(idConnection) = False) Then
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) And _
     (GotPacketWarning(idConnection) = False) And (DangerGM(idConnection) = False) And _
     (gtc > lootTimeExpire(idConnection)) Then
    If TibiaVersionLong >= 760 Then ' do not move runes to hand!
      If RuneMakerOptions(idConnection).activated = True Then
        resS = SearchItem(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank))
        cond1 = ((mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2)) Or _
          ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2))
        If (resS.foundcount = 0) Then
          tmpcond = True
        Else
          'tmpcond = False
          tmpcond = True
        End If
        cond2 = (tmpcond = True) And (cond1 = False)
        cond3 = mySoulpoints(idConnection) < RuneMakerOptions(idConnection).secondActionSoulpoints
        If (cond2 Or cond3) Then
            ' can't make rune
            makingRune(idConnection) = False 'not making rune mode
            runemakerMana1(idConnection) = -1
            If RuneMakerOptions(idConnection).autoLogoutOutOfRunes = True Then
              If ReconnectionStage(idConnection) = 0 Then
                frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " did runemaker logout - logged out because : out of runes or soulpoints"
                sCheat = "14"
                SafeCastCheatString "TimerMaker1", idConnection, sCheat
                aRes = GiveServerError("Runemaker logout - logged out because : out of runes or soulpoints", idConnection)
                DoEvents
                frmMain.DoCloseActions idConnection
                DoEvents
              End If
            ' waste mana option
            ElseIf (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              'If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              'End If
            End If
        Else
            If runemakerMana1(idConnection) = -1 Then
              runemakerMana1(idConnection) = gtc + randomNumberBetween(0, RunemakerChaos2)
            End If
            If gtc >= runemakerMana1(idConnection) Then
                If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
                    If ((runeTurn(idConnection) Mod 5) = 0) Then
                  ' make the rune now!
                  aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                  makingRune(idConnection) = False 'all is ok again, not making rune mode
                  runemakerMana1(idConnection) = -1
                    End If
                End If
            End If
        End If
      End If
    Else ' old mode
      If RuneMakerOptions(idConnection).activated = True Then
          resS = SearchItem(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank))
          cond1 = ((mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2)) Or _
            ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2))
          If (resS.foundcount = 0) Then
            tmpcond = True
          Else
            'tmpcond = False
            tmpcond = True
          End If
          cond2 = (tmpcond = True) And (cond1 = False)
          cond3 = mySoulpoints(idConnection) < RuneMakerOptions(idConnection).secondActionSoulpoints
          
          cond4 = (UseRightHand.Value = True) And _
             (Not ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = savedItem(idConnection).t1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = savedItem(idConnection).t2)))
            
          cond5 = (UseLeftHand.Value = True) And _
             (Not ((mySlot(idConnection, SLOT_LEFTHAND).t1 = savedItem(idConnection).t1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = savedItem(idConnection).t2)))
                
          If (cond2 Or cond3) And (Not (cond4 Or cond5)) Then
            ' out of runes or soulpoints
            ' logout option
            makingRune(idConnection) = False
            If RuneMakerOptions(idConnection).autoLogoutOutOfRunes = True Then
              If ReconnectionStage(idConnection) = 0 Then
                frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " did runemaker logout - logged out because : out of runes or soulpoints"
                sCheat = "14"
                SafeCastCheatString "TimerMaker2", idConnection, sCheat

                aRes = GiveServerError("Runemaker logout - logged out because : out of runes or soulpoints", idConnection)
                DoEvents
                frmMain.DoCloseActions idConnection
                DoEvents
              End If
            ' waste mana option
            ElseIf (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              End If
            End If
          Else
            If runemakerMana1(idConnection) = -1 Then
              runemakerMana1(idConnection) = gtc + randomNumberBetween(0, RunemakerChaos2)
            End If
            If gtc >= runemakerMana1(idConnection) Then
          
            If UseRightHand.Value = True Then
             ' right hand
            If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
              If mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1 And mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2 Then
                If ((runeTurn(idConnection) Mod 5) = 0) Then
                ' make the rune now!
               ' MODIFIFIED IN 9.35 to allow exiva testsound
                aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                End If
              ElseIf ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = &H0) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = &H0)) Then
                'move blank rune to right hand
                initialRuneBackpack(idConnection) = resS.bpID
                aRes = MoveItemToRightHand(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank), 0, resS.bpID, resS.slotID, False)
              Else
                makingRune(idConnection) = True
  
                aRes = SaveHand(idConnection, True, CByte(SLOT_RIGHTHAND), initialRuneBackpack(idConnection))
                If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
                End If
                DoEvents
                
              End If
            ElseIf ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = 0) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = 0)) Then
                aRes = MoveItemToRightHand(idConnection, savedItem(idConnection).t1, savedItem(idConnection).t2, savedItem(idConnection).t3, 0, 0, True)
                
            ElseIf (Not (mySlot(idConnection, SLOT_RIGHTHAND).t1 = savedItem(idConnection).t1 And mySlot(idConnection, SLOT_RIGHTHAND).t2 = savedItem(idConnection).t2)) Then
              ' put made rune in backpack
              aRes = SaveHand(idConnection, False, CByte(SLOT_RIGHTHAND), initialRuneBackpack(idConnection))
              If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
              End If
              DoEvents
            Else
              makingRune(idConnection) = False 'all is ok again, not making rune mode
              runemakerMana1(idConnection) = -1
            End If
            
            
            Else
              'left hand
            If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
              If mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1 And mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2 Then
                If ((runeTurn(idConnection) Mod 5) = 0) Then
                ' make the rune now!
               ' MODIFIFIED IN 9.35 to allow exiva testsound
                aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                End If
              ElseIf ((mySlot(idConnection, SLOT_LEFTHAND).t1 = &H0) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = &H0)) Then
               'move blank rune to left hand
                initialRuneBackpack(idConnection) = resS.bpID
                aRes = MoveItemToLeftHand(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank), 0, resS.bpID, resS.slotID, False)
                
              Else
               makingRune(idConnection) = True

                aRes = SaveHand(idConnection, True, CByte(SLOT_LEFTHAND), initialRuneBackpack(idConnection))
                If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
                End If
                DoEvents
              End If
            ElseIf ((mySlot(idConnection, SLOT_LEFTHAND).t1 = 0) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = 0)) Then
                aRes = MoveItemToLeftHand(idConnection, savedItem(idConnection).t1, savedItem(idConnection).t2, savedItem(idConnection).t3, 0, 0, True)
                
            ElseIf (Not (mySlot(idConnection, SLOT_LEFTHAND).t1 = savedItem(idConnection).t1 And mySlot(idConnection, SLOT_LEFTHAND).t2 = savedItem(idConnection).t2)) Then
              ' put made rune in backpack
              aRes = SaveHand(idConnection, False, CByte(SLOT_LEFTHAND), initialRuneBackpack(idConnection))
              If (aRes = -1) Then
                makingRune(idConnection) = False
                runemakerMana1(idConnection) = -1
              End If
              DoEvents
            Else
              makingRune(idConnection) = False 'all is ok again, not making rune mode
              runemakerMana1(idConnection) = -1
            End If
            
            End If
            
            End If
          End If
      Else
            If (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              End If
            End If
        
      End If
      
    End If ' tibia version
    Else
    End If
    End If ' eat done=false
  Next idConnection
  
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Warning: connection fail during the runemaker function - ignoring"
End Sub

Private Sub TimerPM2_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      
      If (RuneMakerOptions(idConnection).msgSound = True) And PlayPMSound = True Then
            DirectX_PlaySound 4 'play sound PM
            PlayPMSound = False 'stop alarm
      End If
      
          If (RuneMakerOptions(idConnection).autoPM2 = True) And (PlayMsgSound = True) Then
            If InStr(1, var_lastmsg(idConnection), RuneMakerOptions(idConnection).beeploot, vbTextCompare) > 0 Then
            DirectX_PlaySound 4 'play sound when appear msg, usually for loot
            'PlayMsgSound = False ' stop alarm
            var_lastmsg(idConnection) = "" 'reset lastmsg var
            End If
          End If
      
      If (cavebotEnabled(idConnection) = True) And (DangerPK(idConnection) = True) And (chkcd.Value = 1) Then
            DirectX_PlaySound 4 'play sound DANGER
      End If
      
    End If
    Next idConnection
        
End Sub

Private Sub Timerpmax_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autopmax = True) Then
                'aqui começa
                StartPush2 (idConnection)
      End If
            If (RuneMakerOptions(idConnection).autopmax = False) Then
                RemoveSpamOrder idConnection, 2
            End If
    End If
    Next idConnection
End Sub

Private Sub TimerSdt_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoSdt = True) Then
                aRes = ExecuteInTibia("sdmax1", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub timerSS_Timer()
    timerSS.enabled = False
    GetScreenshot frmScreenshot, getScreenshotname()
    DoEvents
End Sub







Private Sub Timerssa_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autossa = True) Then
                aRes = ExecuteInTibia("exiva #09 0C 02", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub Timerssap_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim percent As Long
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       percent = 100 * ((myHP(idConnection) / myMaxHP(idConnection)))
      If (RuneMakerOptions(idConnection).autossap = True) And (percent < RuneMakerOptions(idConnection).thirdActionText) Then
                aRes = ExecuteInTibia("exiva #09 0C 02", idConnection, True)
            End If
      End If
    Next idConnection
    
        For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
If chkActivate.Value = 0 Then
    RuneMakerOptions(runemakerIDselected).activated = True
End If
      End If
    Next idConnection
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
If RuneMakerOptions(runemakerIDselected).activated = True Then
    chkActivate.Value = 1
End If
      End If
    Next idConnection
    
End Sub

Private Sub Timertar_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autotar = True) Then
            aRes = ExecuteInTibia("holdtarget2", idConnection, True)
            End If
    End If
    Next idConnection
'    For idConnection = 1 To MAXCLIENTS
'    If (GameConnected(idConnection) = True) And _
'       (sentWelcome(idConnection) = True) Then
'        aRes = ExecuteInTibia("holdtarget", idConnection, True)
'    End If
'    Next idConnection
End Sub

Private Sub Timertestera2_Timer()
Dim aRes As Integer
Dim idConnection As Integer
DIV.GetDeviceStateKeyboard KeyB
'controle = False

For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
     
      If (RuneMakerOptions(idConnection).autodk = True) Then
       
        If KeyB.key(71) > 0 Then  ' NE
        aRes = ExecuteInFocusedTibia("exiva > 6D")
        End If

        If KeyB.key(73) > 0 Then  'ND
        aRes = ExecuteInFocusedTibia("exiva > 6A")
        End If

        If KeyB.key(79) > 0 Then  'SE
        aRes = ExecuteInFocusedTibia("exiva > 6C")
        End If

        If KeyB.key(81) > 0 Then  'SD
        aRes = ExecuteInFocusedTibia("exiva > 6B")
        End If
  
      End If
     
    End If
Next idConnection

End Sub

Private Sub TimerUtamo_Timer()
Dim aRes As Long
Dim sutm As String
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoUtamo = True) And (GetStatusBit(idConnection, 4) = 0) Then
                aRes = ExecuteInTibia("utamo vita", idConnection, True)
            End If
      End If
    Next idConnection
End Sub











Private Sub Timerxray_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer
    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoxray = True) Then
                aRes = ExecuteInTibia("exiva all", idConnection, True)
            End If
      End If
    Next idConnection
End Sub

Private Sub txtAction1_Validate(Cancel As Boolean)
If runemakerIDselected > 0 Then
  RuneMakerOptions(runemakerIDselected).firstActionText = txtAction1.Text
End If
End Sub
Private Sub txtAction2_Validate(Cancel As Boolean)
If runemakerIDselected > 0 Then
  RuneMakerOptions(runemakerIDselected).secondActionText = txtAction2.Text
End If
End Sub









Private Sub txtbeep_Validate(Cancel As Boolean)
If runemakerIDselected > 0 Then
  RuneMakerOptions(runemakerIDselected).beeploot = txtbeep.Text
End If
End Sub

Private Sub txtLowMana_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtLowMana.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).LowMana = lonN
  Else
    txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
    RuneMakerOptions(runemakerIDselected).LowMana = RuneMakerOptions_LowMana_default
  End If
  End If
  Exit Sub
gotError:
  txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
  RuneMakerOptions(runemakerIDselected).LowMana = RuneMakerOptions_LowMana_default
End Sub

Private Sub txtManaAction1_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtManaAction1.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).firstActionMana = lonN
  Else
    txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
    RuneMakerOptions(runemakerIDselected).firstActionMana = RuneMakerOptions_firstActionMana_default
  End If
  End If
  Exit Sub
gotError:
  txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
  RuneMakerOptions(runemakerIDselected).firstActionMana = RuneMakerOptions_firstActionMana_default
End Sub
Private Sub txtManaAction2_Validate(Cancel As Boolean)
 Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtManaAction2.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).secondActionMana = lonN
  Else
    txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
    RuneMakerOptions(runemakerIDselected).secondActionMana = RuneMakerOptions_secondActionMana_default
  End If
  End If
  Exit Sub
gotError:
  txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
  RuneMakerOptions(runemakerIDselected).secondActionMana = RuneMakerOptions_secondActionMana_default
End Sub

Private Sub txtPickupID_Change()
  Dim res As TypePairOfBytes
  If runemakerIDselected > 0 Then
    res = safeConvertStringToPairOfBytes(txtPickupID.Text)
    TrainerOptions(runemakerIDselected).spearID_b1 = res.b1
    TrainerOptions(runemakerIDselected).spearID_b2 = res.b2
  End If
End Sub

Private Sub txtSlotRefill_Change(index As Integer)
  Dim res As TypePairOfBytes
  Dim strTmp As String
  'txtSlotRefill(Index).Text = Combo1.List(Combo1.ListIndex)
  strTmp = txtSlotRefill(index).Text
  'strTmp = Combo1.List(Combo1.ListIndex)
  If ((trainerIDselected > 0) And (index > 0)) Then
    res = safeConvertStringToPairOfBytes(strTmp)
    TrainerOptions(trainerIDselected).PlayerSlots(index).itemID_b1 = res.b1
    TrainerOptions(trainerIDselected).PlayerSlots(index).itemID_b2 = res.b2
  End If
End Sub

Private Sub txtSoulAction2_Validate(Cancel As Boolean)
Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lonN = CLng(txtSoulAction2.Text)
  If runemakerIDselected > 0 Then
  If lonN >= 0 Then
    RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = lonN
  Else
    txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
    RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default
  End If
  End If
  Exit Sub
gotError:
  txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
  RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default
End Sub

Private Sub UseLeftHand_Click()
  Dim aRes As Long
  Dim i As Integer
    #If FinalMode Then
  On Error GoTo gotError
  #End If
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True And GotPacketWarning(i) = False And RuneMakerOptions(i).activated = True Then
      DisableAll i
      aRes = GiveGMmessage(i, "Runemaker have been disabled because the change of hand option. You should reactivate it now.", "Blackd")
      DoEvents
    End If
  Next i
  Exit Sub
gotError:
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error at UseLeftHand_Click"
End Sub

Private Sub UseRightHand_Click()
  Dim aRes As Long
  Dim i As Integer
    #If FinalMode Then
  On Error GoTo gotError
  #End If
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True And GotPacketWarning(i) = False And RuneMakerOptions(i).activated = True Then
      DisableAll i
      aRes = GiveGMmessage(i, "Runemaker have been disabled because the change of hand option. You should reactivate it now.", "Blackd")
      DoEvents
    End If
  Next i
  Exit Sub
gotError:
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error at UseRightHand_Click"
End Sub






