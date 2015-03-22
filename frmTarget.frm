VERSION 5.00
Begin VB.Form frmTarget 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Targeting"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   Icon            =   "frmTarget.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetExoriVis 
      BackColor       =   &H80000014&
      Caption         =   "Exori Vis"
      Height          =   320
      Left            =   3720
      TabIndex        =   135
      ToolTipText     =   "Kill monster with exori vis, also forces standing in front"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetSDkill 
      BackColor       =   &H80000014&
      Caption         =   "SD Attack"
      Height          =   320
      Left            =   3720
      TabIndex        =   132
      ToolTipText     =   "Set the cavebot to kill it with SD runes"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetMeleeKill 
      BackColor       =   &H80000014&
      Caption         =   "Melee Attack"
      Height          =   320
      Left            =   3720
      TabIndex        =   137
      ToolTipText     =   "Allows melee kill of this creature"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtSetMeleeKill 
      Height          =   320
      Left            =   4920
      TabIndex        =   139
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtSetHmmKill 
      Height          =   320
      Left            =   4920
      TabIndex        =   138
      Text            =   "scarab"
      ToolTipText     =   "Enter creature name"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSetHmmKill 
      BackColor       =   &H80000014&
      Caption         =   "HMM rune"
      Height          =   320
      Left            =   3720
      TabIndex        =   136
      ToolTipText     =   "Set the cavebot to kill it with HMM runes"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtExori 
      Height          =   320
      Left            =   4920
      TabIndex        =   134
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtSetSDKill 
      Height          =   320
      Left            =   4920
      TabIndex        =   133
      Text            =   "demon"
      ToolTipText     =   "Enter creature name"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtMort 
      Height          =   320
      Left            =   9240
      TabIndex        =   131
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdSetExoriMort 
      BackColor       =   &H80000014&
      Caption         =   "Exori Mort"
      Height          =   320
      Left            =   8040
      TabIndex        =   130
      ToolTipText     =   "Kill monster with exori mort, also forces standing in front"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtAvoid 
      Height          =   320
      Left            =   4800
      TabIndex        =   23
      Text            =   "dragon"
      ToolTipText     =   "Enter creature name"
      Top             =   4800
      Width           =   795
   End
   Begin VB.TextBox txtPriority2 
      Height          =   320
      Left            =   5400
      MaxLength       =   7
      TabIndex        =   18
      Text            =   "+1"
      ToolTipText     =   "positive values = more priority than default ; negative values = less priority than default"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtPriority1 
      Height          =   320
      Left            =   4200
      TabIndex        =   19
      Text            =   "necromancer"
      ToolTipText     =   "Enter creature name"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetPriority 
      BackColor       =   &H80000014&
      Caption         =   "Priority"
      Height          =   320
      Left            =   3240
      TabIndex        =   20
      ToolTipText     =   "set more priority in some monsters. Default = 0 ; higher value = more priority"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSetFriendly 
      BackColor       =   &H80000014&
      Caption         =   "Respect"
      Height          =   320
      Left            =   5400
      TabIndex        =   50
      ToolTipText     =   "Avoid attacking others creatures"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSetAny 
      BackColor       =   &H80000014&
      Caption         =   "Others"
      Height          =   320
      Left            =   7320
      TabIndex        =   51
      ToolTipText     =   "Attack any creature (rookgard - nonpvps)"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetAvoidFront 
      BackColor       =   &H80000014&
      Caption         =   "Avoid Wave (circle)"
      Height          =   320
      Left            =   3240
      TabIndex        =   25
      ToolTipText     =   "Avoid front of monster"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox chkRunt 
      BackColor       =   &H80000018&
      Caption         =   "Run Targeting"
      Height          =   255
      Left            =   240
      TabIndex        =   93
      ToolTipText     =   "Activate cavebot for this character"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ListBox lstScript2 
      Height          =   1620
      Left            =   240
      TabIndex        =   92
      Top             =   960
      Width           =   3135
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   240
      TabIndex        =   91
      Text            =   "-"
      ToolTipText     =   "Select one of your connected characters"
      Top             =   540
      Width           =   3135
   End
   Begin VB.TextBox txtGotoScriptLine 
      Height          =   375
      Left            =   7500
      TabIndex        =   90
      Text            =   "0"
      ToolTipText     =   "Jump to this script line"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtOnDangerGoto 
      Height          =   375
      Left            =   6360
      TabIndex        =   89
      Text            =   "0"
      ToolTipText     =   "jump to this script line"
      Top             =   9660
      Width           =   615
   End
   Begin VB.TextBox txtWait 
      Height          =   375
      Left            =   5460
      TabIndex        =   88
      Text            =   "10"
      ToolTipText     =   "time in seconds"
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox txtSayMessage 
      Height          =   375
      Left            =   5280
      TabIndex        =   87
      Text            =   "message"
      ToolTipText     =   "Say this message at this script point"
      Top             =   8100
      Width           =   1335
   End
   Begin VB.TextBox txtSetLoot 
      Height          =   320
      Left            =   4440
      TabIndex        =   86
      Text            =   "D7 0B"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: gold"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtIfOne_Ammount 
      Height          =   375
      Left            =   6780
      TabIndex        =   85
      Text            =   "1000"
      ToolTipText     =   "at least this ammount to validate condition"
      Top             =   12300
      Width           =   495
   End
   Begin VB.TextBox txtIfOne_Goto 
      Height          =   375
      Left            =   0
      TabIndex        =   84
      Text            =   "0"
      ToolTipText     =   "if condition is validated then jump here"
      Top             =   14280
      Width           =   1575
   End
   Begin VB.TextBox txtIfOne_Item 
      Height          =   375
      Left            =   6060
      TabIndex        =   83
      Text            =   "D7 0B"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: gold"
      Top             =   12300
      Width           =   615
   End
   Begin VB.TextBox txtIfTwo_Ammount 
      Height          =   375
      Left            =   6780
      TabIndex        =   82
      Text            =   "5"
      ToolTipText     =   "this ammount or less to validate condition"
      Top             =   12660
      Width           =   495
   End
   Begin VB.TextBox txtIfTwo_Goto 
      Height          =   375
      Left            =   0
      TabIndex        =   81
      Text            =   "0"
      ToolTipText     =   "if condition is validated then jump here"
      Top             =   14640
      Width           =   1575
   End
   Begin VB.TextBox txtIfTwo_Item 
      Height          =   375
      Left            =   6060
      TabIndex        =   80
      Text            =   "58 0C"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: UH"
      Top             =   12660
      Width           =   615
   End
   Begin VB.TextBox txtEdit 
      Height          =   375
      Left            =   240
      TabIndex        =   79
      Top             =   3180
      Width           =   3135
   End
   Begin VB.CommandButton cmdDeleteSelected 
      BackColor       =   &H80000014&
      Caption         =   "Del"
      Height          =   255
      Left            =   2760
      TabIndex        =   78
      ToolTipText     =   "Deletes current selected item in the list box"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdLoadScript 
      BackColor       =   &H80000014&
      Caption         =   "Load"
      Height          =   255
      Left            =   1920
      TabIndex        =   77
      ToolTipText     =   "Loads from given file"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveScript 
      BackColor       =   &H80000014&
      Caption         =   "Save"
      Height          =   255
      Left            =   1080
      TabIndex        =   76
      ToolTipText     =   "Saves to given file"
      Top             =   4800
      Width           =   795
   End
   Begin VB.TextBox txtFishTimes 
      Height          =   375
      Left            =   8580
      TabIndex        =   75
      Text            =   "10"
      ToolTipText     =   "aprox number of casts desired"
      Top             =   9600
      Width           =   615
   End
   Begin VB.TextBox txtMs 
      Height          =   285
      Left            =   120
      TabIndex        =   74
      Text            =   "200"
      Top             =   7980
      Width           =   615
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H80000014&
      Caption         =   "Change"
      Height          =   375
      Left            =   2100
      TabIndex        =   73
      ToolTipText     =   "Change global timer"
      Top             =   7980
      Width           =   735
   End
   Begin VB.CheckBox chkChangePkHeal 
      BackColor       =   &H00000000&
      Caption         =   "Change % autoheal at PK to"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   72
      Top             =   10080
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.HScrollBar scrollPkHeal 
      Height          =   255
      Left            =   600
      Max             =   100
      TabIndex        =   71
      Top             =   9480
      Value           =   75
      Width           =   2175
   End
   Begin VB.CommandButton cmdReload 
      BackColor       =   &H80000014&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   240
      TabIndex        =   70
      Top             =   5160
      Width           =   795
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      Left            =   240
      TabIndex        =   69
      Text            =   "default.txt"
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdGotoScriptLine 
      BackColor       =   &H80000018&
      Caption         =   "Loop"
      Height          =   320
      Left            =   3240
      TabIndex        =   68
      ToolTipText     =   "When script read this command, it will jump to given line"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdWait 
      BackColor       =   &H80000014&
      Caption         =   "Wait"
      Height          =   375
      Left            =   4680
      TabIndex        =   67
      ToolTipText     =   "Wait some seconds at this script point"
      Top             =   7560
      Width           =   795
   End
   Begin VB.CommandButton cmdSayMessage 
      BackColor       =   &H80000014&
      Caption         =   "say in Default"
      Height          =   375
      Left            =   2880
      TabIndex        =   66
      ToolTipText     =   "Always say this message at this script point"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0000C000&
      Caption         =   "closeConnection"
      Height          =   375
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "close conection for this client"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdIfOne 
      BackColor       =   &H0000C000&
      Caption         =   "IfEnoughItemsGoto"
      Height          =   375
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Condition. Example: if gold >= 1000 then go to house and drop loot"
      Top             =   12300
      Width           =   2775
   End
   Begin VB.CommandButton cmdIfTwo 
      BackColor       =   &H0000C000&
      Caption         =   "IfFewItemsGoto"
      Height          =   375
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Condition. Example: if count(UHs) < 5  go to safe and logout"
      Top             =   12660
      Width           =   2775
   End
   Begin VB.CommandButton cmdDropLootOnGround 
      BackColor       =   &H0000C000&
      Caption         =   "dropLootOnGround"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Drop all loot of your containers on ground (house or guildhall)"
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton cmdPutLootOnDepot 
      BackColor       =   &H0000C000&
      Caption         =   "putLootOnDepot"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Put your loot inside a depot"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdFish 
      BackColor       =   &H0000C000&
      Caption         =   "fishX"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Fish X times here"
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStackItems 
      BackColor       =   &H80000014&
      Caption         =   "stackItems"
      Height          =   375
      Left            =   2880
      TabIndex        =   59
      ToolTipText     =   "Do all possible stacking "
      Top             =   8580
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetNoFollow 
      BackColor       =   &H0000C000&
      Caption         =   "setNoFollow"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Set mode don't follow targets"
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetFollow 
      BackColor       =   &H0000C000&
      Caption         =   "setFollow"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Set mode follow targets"
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H80000014&
      Caption         =   "Walk"
      Height          =   375
      Left            =   2880
      TabIndex        =   56
      ToolTipText     =   "Move to this position"
      Top             =   7560
      Width           =   795
   End
   Begin VB.CommandButton cmdUseItem 
      BackColor       =   &H80000014&
      Caption         =   "Ladder"
      Height          =   375
      Left            =   3780
      TabIndex        =   55
      ToolTipText     =   "Use an item like a ladder or a switch"
      Top             =   7560
      Width           =   795
   End
   Begin VB.CommandButton cmdSetLootOn 
      BackColor       =   &H0000C000&
      Caption         =   "setLootOn"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Change loot mode"
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton cmdSetLootOff 
      BackColor       =   &H0000C000&
      Caption         =   "setLootOff"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Change loot mode"
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton cmdOnGMcloseConnection 
      BackColor       =   &H0000C000&
      Caption         =   "onGMcloseConnection"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "disconnects you when a gm comes near you"
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton cmdOnGMpause 
      BackColor       =   &H0000C000&
      Caption         =   "onGMpause"
      Height          =   375
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "If you get a gm pop , pause all automatic functions - Enabled by default"
      Top             =   9540
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetVery 
      BackColor       =   &H0000C000&
      Caption         =   "setVeryFriendly"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Avoid attack anything whenever a player is on screen"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetLoot 
      BackColor       =   &H80000014&
      Caption         =   "Loot ID (hexa)"
      Height          =   320
      Left            =   3240
      TabIndex        =   47
      ToolTipText     =   "Allow looting this. Example: Gold"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOnDangerGoto 
      BackColor       =   &H0000C000&
      Caption         =   "onDangerGoto"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "If you get attacked by other creature then jump to this script line"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdResetLoot 
      BackColor       =   &H0000C000&
      Caption         =   "resetLoot"
      Height          =   375
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "resets the list of lootable items"
      Top             =   9180
      Width           =   975
   End
   Begin VB.CommandButton cmdOnTrapGiveAlarm 
      BackColor       =   &H0000C000&
      Caption         =   "onTrapGiveAlarm"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Will give sound alarm at potential traps"
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdOnPlayerPause 
      BackColor       =   &H0000C000&
      Caption         =   "onPLAYERpause-"
      Height          =   375
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "If you get a player , pause all automatic functions - you wont even autouh! - DO NOT USE  IF NOT NEAR COMPUTER"
      Top             =   9540
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "IfTrue ("
      Height          =   375
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "If it is true then jump to given line"
      Top             =   11580
      Width           =   735
   End
   Begin VB.TextBox txtThing1 
      Height          =   375
      Left            =   3900
      TabIndex        =   41
      Text            =   "$mymana$"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   11580
      Width           =   855
   End
   Begin VB.ComboBox cmbOperator 
      Height          =   315
      Left            =   4860
      TabIndex        =   40
      Text            =   "#number<=#"
      ToolTipText     =   "Operator"
      Top             =   11580
      Width           =   1455
   End
   Begin VB.TextBox txtThing2 
      Height          =   375
      Left            =   6420
      TabIndex        =   39
      Text            =   "100"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   11580
      Width           =   855
   End
   Begin VB.TextBox txtLineIFTRUE 
      Height          =   375
      Left            =   720
      TabIndex        =   38
      Text            =   "0"
      ToolTipText     =   "Jump to this script line"
      Top             =   13560
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Try alternative path (old mode)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   37
      Top             =   11640
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox txtBlockSec 
      Height          =   285
      Left            =   2400
      TabIndex        =   36
      Text            =   "30000"
      Top             =   12480
      Width           =   735
   End
   Begin VB.CommandButton cmdChangeTimer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   285
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   12480
      Width           =   375
   End
   Begin VB.CommandButton cmdComment 
      BackColor       =   &H0000C000&
      Caption         =   "Comment ( # )"
      Height          =   375
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Comment lines (not executed)"
      Top             =   13020
      Width           =   1695
   End
   Begin VB.TextBox txtComment 
      Height          =   375
      Left            =   4860
      TabIndex        =   33
      Text            =   "script for my favourite dungeon"
      ToolTipText     =   "Text"
      Top             =   13020
      Width           =   2415
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H0000C000&
      Caption         =   "Label ( : )"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "$nlineoflabel:labelname$ translate to its line"
      Top             =   15000
      Width           =   1695
   End
   Begin VB.TextBox txtLabel 
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      Text            =   "labelname"
      ToolTipText     =   "Text"
      Top             =   15000
      Width           =   1935
   End
   Begin VB.CommandButton fastExiva 
      BackColor       =   &H0000C000&
      Caption         =   "fastExiva"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "process a exiva command and instantly jump to next line"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txtFastExivaMessage 
      Height          =   375
      Left            =   5160
      TabIndex        =   29
      Text            =   "_myvariable = 1"
      ToolTipText     =   "Execute this exiva command and jump to next line instantly"
      Top             =   9120
      Width           =   3015
   End
   Begin VB.CommandButton cmdRetryAttacks 
      BackColor       =   &H0000C000&
      Caption         =   "setRetryAttacks"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Attack the monster all the time (DEFAULT)"
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdDontRetryAttacks 
      BackColor       =   &H0000C000&
      Caption         =   "setDontRetryAttacks"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Send attack order once. This might be dangerous if this order is lost."
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdResetKillables 
      BackColor       =   &H0000C000&
      Caption         =   "resetKill"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "reset setMeleeKill and setHmmKill"
      Top             =   8640
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Kill the monsters when you have been blocked more than ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   12000
      Width           =   3255
   End
   Begin VB.HScrollBar scrollExorivis 
      Height          =   255
      Left            =   1320
      Max             =   100
      TabIndex        =   22
      Top             =   9840
      Value           =   75
      Width           =   1455
   End
   Begin VB.CheckBox chkLootProtection 
      BackColor       =   &H00000000&
      Caption         =   "Allow looting when a person is near (if using a friendly mode)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   12840
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CommandButton cmdSetSpellKill 
      BackColor       =   &H80000014&
      Caption         =   "Set Creature"
      Height          =   320
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "set more priority in some monsters. Default = 0 ; higher value = more priority"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtSetSpellKill_Creature 
      Height          =   320
      Left            =   4920
      TabIndex        =   16
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtSetSpellKill_Spell 
      Height          =   320
      Left            =   3720
      TabIndex        =   15
      Text            =   "exori frigo"
      ToolTipText     =   "Enter distance spell"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtSetSpellKill_Dist 
      Height          =   320
      Left            =   5400
      TabIndex        =   14
      Text            =   "3"
      ToolTipText     =   "Enter maximum distance for possible cast"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdSayInTrade 
      BackColor       =   &H80000014&
      Caption         =   "say in Channel"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      ToolTipText     =   "say this message in trade, if trading"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetLootDistance 
      BackColor       =   &H0000C000&
      Caption         =   "setLootDistance"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Change max distance to corpse to be lootable"
      Top             =   9180
      Width           =   1695
   End
   Begin VB.TextBox txtSetLootDistance 
      Height          =   375
      Left            =   8340
      TabIndex        =   11
      Text            =   "3"
      ToolTipText     =   "max distance to the corpse"
      Top             =   9060
      Width           =   615
   End
   Begin VB.CommandButton cmdSetMaxAttackTimeMs 
      BackColor       =   &H0000C000&
      Caption         =   "setMaxAttackTimeMs"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "if you take more time than that to kill target, then ignore it"
      Top             =   9540
      Width           =   1935
   End
   Begin VB.TextBox txtSetMaxAttackTimeMs 
      Height          =   375
      Left            =   3780
      TabIndex        =   9
      Text            =   "40000"
      ToolTipText     =   "if you take more time than that to kill target, then ignore it"
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton cmdSetMaxHit 
      BackColor       =   &H0000C000&
      Caption         =   "setMaxHit"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "If a target hits you more than this then, then ignore it"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtSetMaxHit 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Text            =   "10000"
      ToolTipText     =   "If a target hits you more than this then, then ignore it"
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox txtMs2 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "200"
      Top             =   7980
      Width           =   615
   End
   Begin VB.CommandButton cmdSetChaoticMovesON 
      BackColor       =   &H0000C000&
      Caption         =   "setChaoticMovesON"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "It will avoid repetitive path detection (enabled by default)"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetChaoticMovesOFF 
      BackColor       =   &H0000C000&
      Caption         =   "setChaoticMovesOFF"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Try to always move to exact waypoint"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetBot 
      BackColor       =   &H0000C000&
      Caption         =   "setBot"
      Height          =   375
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "set internal bot variable"
      Top             =   9000
      Width           =   735
   End
   Begin VB.ComboBox cmbSetOperator 
      Height          =   315
      Left            =   6900
      TabIndex        =   2
      Text            =   "LootAll"
      ToolTipText     =   "Bot internal variable"
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox txtSetBotValue 
      Height          =   375
      Left            =   8940
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "value, for booleans 0=FALSE and 1=TRUE"
      Top             =   9000
      Width           =   735
   End
   Begin VB.CommandButton cmdLoadCopyPaste 
      BackColor       =   &H80000014&
      Caption         =   "Edit"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Loads from given file"
      Top             =   4800
      Width           =   795
   End
   Begin VB.Line Line18 
      BorderColor     =   &H80000002&
      X1              =   6840
      X2              =   6840
      Y1              =   4320
      Y2              =   6000
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000002&
      X1              =   3120
      X2              =   6840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000002&
      X1              =   3120
      X2              =   6840
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000002&
      X1              =   3120
      X2              =   3120
      Y1              =   4320
      Y2              =   6000
   End
   Begin VB.Line Line14 
      X1              =   7080
      X2              =   9960
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line13 
      X1              =   10560
      X2              =   10560
      Y1              =   4440
      Y2              =   6120
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000002&
      X1              =   3600
      X2              =   6840
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000002&
      X1              =   6840
      X2              =   6840
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000002&
      X1              =   3600
      X2              =   6840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000002&
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000002&
      X1              =   2880
      X2              =   2880
      Y1              =   4320
      Y2              =   6000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   2880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   2880
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   120
      Y1              =   4320
      Y2              =   6000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000002&
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   3480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   120
      X2              =   3480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000018&
      Caption         =   "creature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   101
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000018&
      Caption         =   "="
      Height          =   255
      Left            =   5280
      TabIndex        =   106
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000018&
      Caption         =   "Select your character :"
      Height          =   255
      Left            =   240
      TabIndex        =   129
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblConfigComands 
      BackColor       =   &H80000018&
      Caption         =   "MONSTER TARGETING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   128
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label lblActions 
      BackColor       =   &H80000018&
      Caption         =   "Desired Action                  Creature"
      Height          =   255
      Left            =   3720
      TabIndex        =   127
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblScriptCommands 
      BackColor       =   &H00000000&
      Caption         =   "Script commands:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3180
      TabIndex        =   126
      Top             =   11100
      Width           =   1815
   End
   Begin VB.Label lblFile 
      BackColor       =   &H00000000&
      Caption         =   "File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   125
      Top             =   8940
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "itemID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6165
      TabIndex        =   124
      Top             =   12060
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6795
      TabIndex        =   123
      Top             =   12060
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   122
      Top             =   14040
      Width           =   375
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000018&
      Caption         =   "ElfBot cavebot !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   1215
      Left            =   240
      TabIndex        =   121
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BackColor       =   &H80000018&
      Caption         =   "Edit line :"
      Height          =   255
      Left            =   240
      TabIndex        =   120
      Top             =   2940
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000018&
      Caption         =   "Set Cavebot speed :"
      Height          =   255
      Left            =   120
      TabIndex        =   119
      Top             =   7620
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000018&
      Caption         =   "to"
      Height          =   375
      Left            =   840
      TabIndex        =   118
      Top             =   7980
      Width           =   255
   End
   Begin VB.Label lblPKhealValue 
      BackColor       =   &H00000000&
      Caption         =   "75 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   117
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   ") Goto"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   116
      Top             =   13560
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "thing1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4020
      TabIndex        =   115
      Top             =   11340
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "thing2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6540
      TabIndex        =   114
      Top             =   11340
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "operator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4860
      TabIndex        =   113
      Top             =   11340
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   112
      Top             =   13320
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000018&
      Caption         =   "Spell Kill"
      Height          =   255
      Left            =   3720
      TabIndex        =   111
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "time(ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   110
      Top             =   12480
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "If blocked by killable monsters not yours:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   109
      Top             =   11400
      Width           =   3375
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000018&
      Caption         =   "LOOTING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   108
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label lblExorivisValue 
      BackColor       =   &H00000000&
      Caption         =   "50 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   107
      Top             =   9840
      Width           =   615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   ","
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   105
      Top             =   9240
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   ","
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   104
      Top             =   9180
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000018&
      Caption         =   "Saving and Loading settings"
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "distance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   420
      TabIndex        =   102
      Top             =   8820
      Width           =   555
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000018&
      Caption         =   "WAYPOINTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   100
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000018&
      Caption         =   "Misc. "
      Height          =   255
      Left            =   3120
      TabIndex        =   99
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label23 
      BackColor       =   &H80000018&
      Caption         =   "Desired Spell Attack        dist."
      Height          =   255
      Left            =   3720
      TabIndex        =   98
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000018&
      Caption         =   "ms"
      Height          =   375
      Left            =   1800
      TabIndex        =   97
      Top             =   7980
      Width           =   375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8865
      TabIndex        =   96
      Top             =   9000
      Width           =   255
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7020
      TabIndex        =   95
      Top             =   8460
      Width           =   3135
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000018&
      Caption         =   "<-use this to finish script"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4560
      TabIndex        =   94
      Top             =   5520
      Width           =   1635
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
