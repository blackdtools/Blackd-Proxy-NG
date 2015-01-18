VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmCavebot 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cavebot"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   ForeColor       =   &H80000018&
   Icon            =   "frmCavebot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton cmdLoadScript 
      Height          =   255
      Left            =   960
      TabIndex        =   163
      Top             =   4860
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      Caption         =   "load"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSaveScript 
      Height          =   255
      Left            =   120
      TabIndex        =   162
      Top             =   4860
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      Caption         =   "save"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdReload 
      Height          =   255
      Left            =   1800
      TabIndex        =   161
      Top             =   4860
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      Caption         =   "refresh"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdLoadCopyPaste 
      Height          =   315
      Left            =   2640
      TabIndex        =   160
      Top             =   4440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Caption         =   "edit"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdDeleteSelected 
      Height          =   195
      Left            =   2640
      TabIndex        =   159
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   344
      Caption         =   "del"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdChange 
      Height          =   315
      Left            =   1920
      TabIndex        =   157
      Top             =   5520
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Caption         =   "change"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdGotoScriptLine 
      Height          =   315
      Left            =   3480
      TabIndex        =   156
      Top             =   4320
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      Caption         =   "Loop (finish script)"
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
      BackColor       =   49152
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSayInTrade 
      Height          =   315
      Left            =   4740
      TabIndex        =   155
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "say in NPC"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSayMessage 
      Height          =   315
      Left            =   3480
      TabIndex        =   154
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "say in Default"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdWait 
      Height          =   315
      Left            =   5280
      TabIndex        =   153
      Top             =   3180
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Wait"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdUseItem 
      Height          =   315
      Left            =   4380
      TabIndex        =   152
      Top             =   3180
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Ladder"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdMove 
      Height          =   315
      Left            =   3480
      TabIndex        =   151
      Top             =   3180
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Walk"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdIdlist 
      Height          =   315
      Left            =   5340
      TabIndex        =   150
      Top             =   2460
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Items ID (hex)"
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
      BackColor       =   8454016
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSetLoot 
      Height          =   315
      Left            =   3480
      TabIndex        =   149
      Top             =   2460
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Loot ID"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSetPriority 
      Height          =   315
      Left            =   3480
      TabIndex        =   148
      Top             =   2100
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Priority"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSetAvoidFront 
      Height          =   315
      Left            =   3480
      TabIndex        =   147
      Top             =   1740
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      Caption         =   "Avoid Wave (circle)"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSetSpellKill 
      Height          =   315
      Left            =   3480
      TabIndex        =   146
      Top             =   1020
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Spell Kill"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSetMeleeKill 
      Height          =   315
      Left            =   3480
      TabIndex        =   145
      Top             =   660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Attack"
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
      BackColor       =   65280
      BorderColor_Hover=   16777215
   End
   Begin VB.CommandButton cmdIdlist2 
      Caption         =   "Items ID (hex)"
      Height          =   375
      Left            =   8460
      TabIndex        =   144
      Top             =   4620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMoveN 
      Caption         =   "Auto Exploring"
      Height          =   375
      Left            =   9360
      TabIndex        =   142
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkRunt 
      Caption         =   "Follow waypoints"
      Height          =   255
      Left            =   7260
      TabIndex        =   141
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton cmdGotoScriptLine2 
      BackColor       =   &H80000018&
      Caption         =   "Loop (finish script)"
      Height          =   315
      Left            =   8340
      MaskColor       =   &H80000018&
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "When script read this command, it will jump to given line"
      Top             =   5100
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoadCopyPaste2 
      BackColor       =   &H80000014&
      Caption         =   "Edit"
      Height          =   315
      Left            =   7500
      TabIndex        =   136
      ToolTipText     =   "Loads from given file"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSetBotValue 
      Height          =   375
      Left            =   8820
      TabIndex        =   134
      Text            =   "1"
      ToolTipText     =   "value, for booleans 0=FALSE and 1=TRUE"
      Top             =   9120
      Width           =   735
   End
   Begin VB.ComboBox cmbSetOperator 
      Height          =   315
      Left            =   6780
      TabIndex        =   133
      Text            =   "LootAll"
      ToolTipText     =   "Bot internal variable"
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetBot 
      BackColor       =   &H0000C000&
      Caption         =   "setBot"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   132
      ToolTipText     =   "set internal bot variable"
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton cmdSetChaoticMovesOFF 
      BackColor       =   &H0000C000&
      Caption         =   "setChaoticMovesOFF"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   "Try to always move to exact waypoint"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetChaoticMovesON 
      BackColor       =   &H0000C000&
      Caption         =   "setChaoticMovesON"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   130
      ToolTipText     =   "It will avoid repetitive path detection (enabled by default)"
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox txtMs2 
      Height          =   285
      Left            =   1020
      TabIndex        =   128
      Text            =   "200"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtSetMaxHit 
      Height          =   375
      Left            =   5520
      TabIndex        =   126
      Text            =   "10000"
      ToolTipText     =   "If a target hits you more than this then, then ignore it"
      Top             =   9000
      Width           =   735
   End
   Begin VB.CommandButton cmdSetMaxHit 
      BackColor       =   &H0000C000&
      Caption         =   "setMaxHit"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   125
      ToolTipText     =   "If a target hits you more than this then, then ignore it"
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox txtSetMaxAttackTimeMs 
      Height          =   375
      Left            =   4200
      TabIndex        =   124
      Text            =   "40000"
      ToolTipText     =   "if you take more time than that to kill target, then ignore it"
      Top             =   10560
      Width           =   735
   End
   Begin VB.CommandButton cmdSetMaxAttackTimeMs 
      BackColor       =   &H0000C000&
      Caption         =   "setMaxAttackTimeMs"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "if you take more time than that to kill target, then ignore it"
      Top             =   10680
      Width           =   1935
   End
   Begin VB.TextBox txtSetLootDistance 
      Height          =   375
      Left            =   8220
      TabIndex        =   122
      Text            =   "3"
      ToolTipText     =   "max distance to the corpse"
      Top             =   9180
      Width           =   615
   End
   Begin VB.CommandButton cmdSetLootDistance 
      BackColor       =   &H0000C000&
      Caption         =   "setLootDistance"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   121
      ToolTipText     =   "Change max distance to corpse to be lootable"
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSayInTrade2 
      BackColor       =   &H80000014&
      Caption         =   "say in NPC"
      Height          =   375
      Left            =   8460
      TabIndex        =   120
      ToolTipText     =   "say this message in trade, if trading"
      Top             =   6540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSetSpellKill_Dist 
      Height          =   315
      Left            =   6480
      TabIndex        =   113
      Text            =   "3"
      ToolTipText     =   "Enter maximum distance for possible cast"
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtSetSpellKill_Spell 
      Height          =   315
      Left            =   5640
      TabIndex        =   112
      Text            =   "exori frigo"
      ToolTipText     =   "Enter distance spell"
      Top             =   1020
      Width           =   795
   End
   Begin VB.TextBox txtSetSpellKill_Creature 
      Height          =   315
      Left            =   4500
      TabIndex        =   111
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetSpellKill2 
      BackColor       =   &H80000014&
      Caption         =   "Spell Kill"
      Height          =   315
      Left            =   8040
      TabIndex        =   110
      ToolTipText     =   "set more priority in some monsters. Default = 0 ; higher value = more priority"
      Top             =   1260
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPriority2 
      Height          =   315
      Left            =   6180
      MaxLength       =   7
      TabIndex        =   109
      Text            =   "+1"
      ToolTipText     =   "positive values = more priority than default ; negative values = less priority than default"
      Top             =   2100
      Width           =   495
   End
   Begin VB.TextBox txtPriority1 
      Height          =   315
      Left            =   4500
      TabIndex        =   107
      Text            =   "necromancer"
      ToolTipText     =   "Enter creature name"
      Top             =   2100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetPriority2 
      BackColor       =   &H80000014&
      Caption         =   "Priority"
      Height          =   375
      Left            =   8160
      TabIndex        =   106
      ToolTipText     =   "set more priority in some monsters. Default = 0 ; higher value = more priority"
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkLootProtection 
      BackColor       =   &H00000000&
      Caption         =   "Allow looting when a person is near (if using a friendly mode)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   105
      Top             =   12960
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CommandButton cmdSetExoriMort 
      BackColor       =   &H80000014&
      Caption         =   "Exori Mort"
      Height          =   375
      Left            =   9060
      TabIndex        =   104
      ToolTipText     =   "Kill monster with exori mort, also forces standing in front"
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox txtMort 
      Height          =   375
      Left            =   10260
      TabIndex        =   103
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   8160
      Width           =   1995
   End
   Begin VB.CommandButton cmdSetSDkill 
      BackColor       =   &H80000014&
      Caption         =   "SD Attack"
      Height          =   375
      Left            =   4020
      TabIndex        =   101
      ToolTipText     =   "Set the cavebot to kill it with SD runes"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txtSetSDKill 
      Height          =   375
      Left            =   5220
      TabIndex        =   102
      Text            =   "demon"
      ToolTipText     =   "Enter creature name"
      Top             =   8400
      Width           =   1995
   End
   Begin VB.HScrollBar scrollExorivis 
      Height          =   255
      Left            =   1200
      Max             =   100
      TabIndex        =   99
      Top             =   9960
      Value           =   75
      Width           =   1455
   End
   Begin VB.TextBox txtAvoid 
      Height          =   315
      Left            =   5640
      TabIndex        =   97
      Text            =   "dragon"
      ToolTipText     =   "Enter creature name"
      Top             =   1740
      Width           =   1035
   End
   Begin VB.TextBox txtExori 
      Height          =   375
      Left            =   5580
      TabIndex        =   96
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   8940
      Width           =   1995
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Kill the monsters when you have been blocked more than ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   80
      Top             =   12120
      Width           =   3255
   End
   Begin VB.CommandButton cmdSetAvoidFront2 
      BackColor       =   &H80000014&
      Caption         =   "Avoid Wave (circle)"
      Height          =   375
      Left            =   8040
      TabIndex        =   95
      ToolTipText     =   "Avoid front of monster"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdResetKillables 
      BackColor       =   &H0000C000&
      Caption         =   "resetKill"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "reset setMeleeKill and setHmmKill"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdSetExoriVis 
      BackColor       =   &H80000014&
      Caption         =   "Exori Vis"
      Height          =   375
      Left            =   4380
      TabIndex        =   93
      ToolTipText     =   "Kill monster with exori vis, also forces standing in front"
      Top             =   8940
      Width           =   1215
   End
   Begin VB.CommandButton cmdDontRetryAttacks 
      BackColor       =   &H0000C000&
      Caption         =   "setDontRetryAttacks"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Send attack order once. This might be dangerous if this order is lost."
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdRetryAttacks 
      BackColor       =   &H0000C000&
      Caption         =   "setRetryAttacks"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "Attack the monster all the time (DEFAULT)"
      Top             =   10440
      Width           =   1695
   End
   Begin VB.TextBox txtFastExivaMessage 
      Height          =   375
      Left            =   5040
      TabIndex        =   90
      Text            =   "_myvariable = 1"
      ToolTipText     =   "Execute this exiva command and jump to next line instantly"
      Top             =   9240
      Width           =   3015
   End
   Begin VB.CommandButton fastExiva 
      BackColor       =   &H0000C000&
      Caption         =   "fastExiva"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "process a exiva command and instantly jump to next line"
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtLabel 
      Height          =   375
      Left            =   1560
      TabIndex        =   88
      Text            =   "labelname"
      ToolTipText     =   "Text"
      Top             =   15120
      Width           =   1935
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H0000C000&
      Caption         =   "Label ( : )"
      Height          =   375
      Left            =   -120
      Style           =   1  'Graphical
      TabIndex        =   87
      ToolTipText     =   "$nlineoflabel:labelname$ translate to its line"
      Top             =   15120
      Width           =   1695
   End
   Begin VB.TextBox txtComment 
      Height          =   375
      Left            =   4740
      TabIndex        =   86
      Text            =   "script for my favourite dungeon"
      ToolTipText     =   "Text"
      Top             =   13140
      Width           =   2415
   End
   Begin VB.CommandButton cmdComment 
      BackColor       =   &H0000C000&
      Caption         =   "Comment ( # )"
      Height          =   375
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Comment lines (not executed)"
      Top             =   13140
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangeTimer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   285
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   12600
      Width           =   375
   End
   Begin VB.TextBox txtBlockSec 
      Height          =   285
      Left            =   2280
      TabIndex        =   81
      Text            =   "30000"
      Top             =   12600
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Try alternative path (old mode)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   79
      Top             =   11760
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox txtLineIFTRUE 
      Height          =   375
      Left            =   600
      TabIndex        =   73
      Text            =   "0"
      ToolTipText     =   "Jump to this script line"
      Top             =   13680
      Width           =   1575
   End
   Begin VB.TextBox txtThing2 
      Height          =   375
      Left            =   6300
      TabIndex        =   71
      Text            =   "100"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   11700
      Width           =   855
   End
   Begin VB.ComboBox cmbOperator 
      Height          =   315
      Left            =   4740
      TabIndex        =   70
      Text            =   "#number<=#"
      ToolTipText     =   "Operator"
      Top             =   11700
      Width           =   1455
   End
   Begin VB.TextBox txtThing1 
      Height          =   375
      Left            =   3780
      TabIndex        =   69
      Text            =   "$mymana$"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   11700
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "IfTrue ("
      Height          =   375
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "If it is true then jump to given line"
      Top             =   11700
      Width           =   735
   End
   Begin VB.CommandButton cmdOnPlayerPause 
      BackColor       =   &H0000C000&
      Caption         =   "onPLAYERpause-"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "If you get a player , pause all automatic functions - you wont even autouh! - DO NOT USE  IF NOT NEAR COMPUTER"
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOnTrapGiveAlarm 
      BackColor       =   &H0000C000&
      Caption         =   "onTrapGiveAlarm"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Will give sound alarm at potential traps"
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdResetLoot 
      BackColor       =   &H0000C000&
      Caption         =   "resetLoot"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "resets the list of lootable items"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdOnDangerGoto 
      BackColor       =   &H0000C000&
      Caption         =   "onDangerGoto"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "If you get attacked by other creature then jump to this script line"
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetHmmKill 
      BackColor       =   &H80000014&
      Caption         =   "HMM rune"
      Height          =   375
      Left            =   9060
      TabIndex        =   9
      ToolTipText     =   "Set the cavebot to kill it with HMM runes"
      Top             =   9060
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetLoot2 
      BackColor       =   &H80000014&
      Caption         =   "Loot ID (hex)"
      Height          =   375
      Left            =   8340
      TabIndex        =   11
      ToolTipText     =   "Allow looting this. Example: Gold"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetMeleeKill2 
      BackColor       =   &H80000014&
      Caption         =   "Attack"
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "Allows melee kill of this creature"
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSetVery 
      BackColor       =   &H0000C000&
      Caption         =   "setVeryFriendly"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Avoid attack anything whenever a player is on screen"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOnGMpause 
      BackColor       =   &H0000C000&
      Caption         =   "onGMpause"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "If you get a gm pop , pause all automatic functions - Enabled by default"
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetFriendly 
      BackColor       =   &H80000014&
      Caption         =   "Respect (dont attack anothers)"
      Height          =   375
      Left            =   4320
      TabIndex        =   57
      ToolTipText     =   "Avoid attacking others creatures"
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdSetAny 
      BackColor       =   &H80000014&
      Caption         =   "Attack any creature (others)"
      Height          =   375
      Left            =   9060
      TabIndex        =   56
      ToolTipText     =   "Attack any creature (rookgard - nonpvps)"
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton cmdOnGMcloseConnection 
      BackColor       =   &H0000C000&
      Caption         =   "onGMcloseConnection"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "disconnects you when a gm comes near you"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetLootOff 
      BackColor       =   &H0000C000&
      Caption         =   "setLootOff"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Change loot mode"
      Top             =   10320
      Width           =   975
   End
   Begin VB.CommandButton cmdSetLootOn 
      BackColor       =   &H0000C000&
      Caption         =   "setLootOn"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Change loot mode"
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton cmdUseItem2 
      BackColor       =   &H80000014&
      Caption         =   "Ladder"
      Height          =   375
      Left            =   8100
      TabIndex        =   16
      ToolTipText     =   "Use an item like a ladder or a switch"
      Top             =   4320
      Visible         =   0   'False
      Width           =   820
   End
   Begin VB.CommandButton cmdMove2 
      BackColor       =   &H80000014&
      Caption         =   "Walk"
      Height          =   375
      Left            =   7980
      TabIndex        =   15
      ToolTipText     =   "Move to this position"
      Top             =   4980
      Visible         =   0   'False
      Width           =   820
   End
   Begin VB.CommandButton cmdSetFollow 
      BackColor       =   &H0000C000&
      Caption         =   "setFollow"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Set mode follow targets"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetNoFollow 
      BackColor       =   &H0000C000&
      Caption         =   "setNoFollow"
      Height          =   375
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Set mode don't follow targets"
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdStackItems 
      BackColor       =   &H80000014&
      Caption         =   "stackItems"
      Height          =   375
      Left            =   7860
      TabIndex        =   44
      ToolTipText     =   "Do all possible stacking "
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdFish 
      BackColor       =   &H0000C000&
      Caption         =   "fishX"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Fish X times here"
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPutLootOnDepot 
      BackColor       =   &H0000C000&
      Caption         =   "putLootOnDepot"
      Height          =   375
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Put your loot inside a depot"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdDropLootOnGround 
      BackColor       =   &H0000C000&
      Caption         =   "dropLootOnGround"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Drop all loot of your containers on ground (house or guildhall)"
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdIfTwo 
      BackColor       =   &H0000C000&
      Caption         =   "IfFewItemsGoto"
      Height          =   375
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Condition. Example: if count(UHs) < 5  go to safe and logout"
      Top             =   12780
      Width           =   2775
   End
   Begin VB.CommandButton cmdIfOne 
      BackColor       =   &H0000C000&
      Caption         =   "IfEnoughItemsGoto"
      Height          =   375
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Condition. Example: if gold >= 1000 then go to house and drop loot"
      Top             =   12420
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0000C000&
      Caption         =   "closeConnection"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "close conection for this client"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdSayMessage2 
      BackColor       =   &H80000014&
      Caption         =   "say in Default"
      Height          =   375
      Left            =   9180
      TabIndex        =   20
      ToolTipText     =   "Always say this message at this script point"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdWait2 
      BackColor       =   &H80000014&
      Caption         =   "Wait"
      Height          =   375
      Left            =   8100
      TabIndex        =   17
      ToolTipText     =   "Wait some seconds at this script point"
      Top             =   3720
      Visible         =   0   'False
      Width           =   820
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      Left            =   120
      TabIndex        =   63
      Text            =   "default.txt"
      Top             =   4440
      Width           =   2475
   End
   Begin VB.CommandButton cmdReload2 
      BackColor       =   &H80000014&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   7500
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.HScrollBar scrollPkHeal 
      Height          =   255
      Left            =   480
      Max             =   100
      TabIndex        =   59
      Top             =   9600
      Value           =   75
      Width           =   2175
   End
   Begin VB.CheckBox chkChangePkHeal 
      BackColor       =   &H00000000&
      Caption         =   "Change % autoheal at PK to"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   58
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdChange2 
      BackColor       =   &H80000014&
      Caption         =   "Change"
      Height          =   255
      Left            =   7920
      TabIndex        =   52
      ToolTipText     =   "Change global timer"
      Top             =   5700
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtMs 
      Height          =   285
      Left            =   120
      TabIndex        =   51
      Text            =   "200"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtFishTimes 
      Height          =   375
      Left            =   8460
      TabIndex        =   46
      Text            =   "10"
      ToolTipText     =   "aprox number of casts desired"
      Top             =   9720
      Width           =   615
   End
   Begin VB.Timer TimerScript 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   480
   End
   Begin VB.CommandButton cmdSaveScript2 
      BackColor       =   &H80000014&
      Caption         =   "Save"
      Height          =   255
      Left            =   7500
      TabIndex        =   6
      ToolTipText     =   "Saves to given file"
      Top             =   3540
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdLoadScript2 
      BackColor       =   &H80000014&
      Caption         =   "Load"
      Height          =   255
      Left            =   7500
      TabIndex        =   5
      ToolTipText     =   "Loads from given file"
      Top             =   3300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdDeleteSelected2 
      BackColor       =   &H80000014&
      Caption         =   "Del"
      Height          =   255
      Left            =   7500
      TabIndex        =   4
      ToolTipText     =   "Deletes current selected item in the list box"
      Top             =   4140
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtIfTwo_Item 
      Height          =   375
      Left            =   5940
      TabIndex        =   31
      Text            =   "58 0C"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: UH"
      Top             =   12780
      Width           =   615
   End
   Begin VB.TextBox txtIfTwo_Goto 
      Height          =   375
      Left            =   -120
      TabIndex        =   33
      Text            =   "0"
      ToolTipText     =   "if condition is validated then jump here"
      Top             =   14760
      Width           =   1575
   End
   Begin VB.TextBox txtIfTwo_Ammount 
      Height          =   375
      Left            =   6660
      TabIndex        =   32
      Text            =   "5"
      ToolTipText     =   "this ammount or less to validate condition"
      Top             =   12780
      Width           =   495
   End
   Begin VB.TextBox txtIfOne_Item 
      Height          =   375
      Left            =   5940
      TabIndex        =   27
      Text            =   "D7 0B"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: gold"
      Top             =   12420
      Width           =   615
   End
   Begin VB.TextBox txtIfOne_Goto 
      Height          =   375
      Left            =   -120
      TabIndex        =   29
      Text            =   "0"
      ToolTipText     =   "if condition is validated then jump here"
      Top             =   14400
      Width           =   1575
   End
   Begin VB.TextBox txtIfOne_Ammount 
      Height          =   375
      Left            =   6660
      TabIndex        =   28
      Text            =   "1000"
      ToolTipText     =   "at least this ammount to validate condition"
      Top             =   12420
      Width           =   495
   End
   Begin VB.TextBox txtSetLoot 
      Height          =   315
      Left            =   4500
      TabIndex        =   12
      Text            =   "D7 0B"
      ToolTipText     =   "Get tileIDs with the tool module. The example is: gold"
      Top             =   2460
      Width           =   735
   End
   Begin VB.TextBox txtSetHmmKill 
      Height          =   375
      Left            =   10260
      TabIndex        =   10
      Text            =   "scarab"
      ToolTipText     =   "Enter creature name"
      Top             =   8580
      Width           =   1995
   End
   Begin VB.TextBox txtSetMeleeKill 
      Height          =   315
      Left            =   4500
      TabIndex        =   8
      Text            =   "larva"
      ToolTipText     =   "Enter creature name"
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox txtSayMessage 
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Text            =   "message"
      ToolTipText     =   "Say this message at this script point"
      Top             =   3900
      Width           =   2475
   End
   Begin VB.TextBox txtWait 
      Height          =   315
      Left            =   6120
      TabIndex        =   18
      Text            =   "10"
      ToolTipText     =   "time in seconds"
      Top             =   3180
      Width           =   555
   End
   Begin VB.TextBox txtOnDangerGoto 
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Text            =   "0"
      ToolTipText     =   "jump to this script line"
      Top             =   9780
      Width           =   615
   End
   Begin VB.TextBox txtGotoScriptLine 
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "Jump to this script line"
      Top             =   9360
      Width           =   495
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "-"
      ToolTipText     =   "Select one of your connected characters"
      Top             =   420
      Width           =   3135
   End
   Begin VB.ListBox lstScript 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CheckBox chkEnabled 
      BackColor       =   &H80000012&
      Caption         =   "Follow waypoints"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Activate cavebot for this character"
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      Caption         =   "Saving and Loading settings"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   158
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblActions 
      BackColor       =   &H80000012&
      Caption         =   "Desired Action     Creature"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Label30 
      BackColor       =   &H80000008&
      Caption         =   "Set Target and Loot first"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   140
      Top             =   5340
      Width           =   2775
   End
   Begin VB.Label Label28 
      BackColor       =   &H80000007&
      Caption         =   "Scripts for Blackd will work here"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   139
      Top             =   5580
      Width           =   2595
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8340
      TabIndex        =   138
      Top             =   6840
      Width           =   1635
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
      Left            =   6900
      TabIndex        =   137
      Top             =   8580
      Width           =   3135
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8745
      TabIndex        =   135
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000012&
      Caption         =   "ms"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1680
      TabIndex        =   129
      Top             =   5580
      Width           =   375
   End
   Begin VB.Label Label23 
      BackColor       =   &H80000017&
      Caption         =   "Spell        dist."
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   5880
      TabIndex        =   127
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000017&
      Caption         =   "Misc."
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3480
      TabIndex        =   119
      Top             =   1500
      Width           =   555
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000017&
      Caption         =   "Emplacement"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3480
      TabIndex        =   118
      Top             =   2940
      Width           =   1275
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
      Left            =   8760
      TabIndex        =   117
      Top             =   7800
      Width           =   495
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
      Left            =   300
      TabIndex        =   116
      Top             =   8940
      Width           =   555
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   ","
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   115
      Top             =   9300
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   ","
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   114
      Top             =   9360
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "="
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   6000
      TabIndex        =   108
      Top             =   2220
      Width           =   255
   End
   Begin VB.Label lblExorivisValue 
      BackColor       =   &H00000000&
      Caption         =   "50 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   100
      Top             =   9960
      Width           =   615
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
      Left            =   11400
      TabIndex        =   98
      Top             =   7500
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "If blocked by killable monsters not yours:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   83
      Top             =   11520
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "time(ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   82
      Top             =   12600
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000018&
      Caption         =   "Spell Kill"
      Height          =   255
      Left            =   7620
      TabIndex        =   78
      Top             =   8520
      Width           =   2295
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
      Left            =   600
      TabIndex        =   77
      Top             =   13440
      Width           =   375
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
      Left            =   4740
      TabIndex        =   76
      Top             =   11460
      Width           =   735
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
      Left            =   6420
      TabIndex        =   75
      Top             =   11460
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
      Left            =   3900
      TabIndex        =   74
      Top             =   11460
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   ") Goto"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -120
      TabIndex        =   72
      Top             =   13680
      Width           =   495
   End
   Begin VB.Label lblPKhealValue 
      BackColor       =   &H00000000&
      Caption         =   "75 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   60
      Top             =   9720
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "to"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   780
      TabIndex        =   50
      Top             =   5580
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Set Cavebot speed :"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblEdit 
      BackColor       =   &H80000017&
      Caption         =   "Edit line :"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3300
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000017&
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
      Left            =   120
      TabIndex        =   45
      Top             =   3000
      Width           =   3135
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
      Left            =   -120
      TabIndex        =   42
      Top             =   14160
      Width           =   375
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
      Left            =   6675
      TabIndex        =   41
      Top             =   12180
      Width           =   615
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
      Left            =   6045
      TabIndex        =   40
      Top             =   12180
      Width           =   495
   End
   Begin VB.Label lblFile 
      BackColor       =   &H00000000&
      Caption         =   "File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   39
      Top             =   9060
      Width           =   975
   End
   Begin VB.Label lblScriptCommands 
      BackColor       =   &H00000000&
      Caption         =   "Script commands:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3060
      TabIndex        =   37
      Top             =   11220
      Width           =   1815
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
      Left            =   6600
      TabIndex        =   35
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000012&
      Caption         =   "Select your character :"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label27 
      BackColor       =   &H80000018&
      Caption         =   "Targeting:"
      Height          =   255
      Left            =   8340
      TabIndex        =   143
      Top             =   660
      Width           =   915
   End
End
Attribute VB_Name = "frmCavebot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit




Public Sub UpdateValues()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim i As Long
  Dim limLines As Long
  lstScript.Clear
  If cavebotIDselected = 0 Then
     If chkEnabled.Value = 1 Then
       avoidC = True
       chkEnabled.Value = 0
       avoidC = False
     End If
  Else
     If cavebotEnabled(cavebotIDselected) = True Then
       If chkEnabled.Value = 0 Then
         avoidC = True
         chkEnabled.Value = 1
         avoidC = False
       End If
              'aq tentativa chk dos form trainer
        
     Else
       If chkEnabled.Value = 1 Then
         avoidC = True
         chkEnabled.Value = 0
         avoidC = False
       End If
     End If
     limLines = cavebotLenght(cavebotIDselected) - 1
     For i = 0 To limLines
       lstScript.AddItem GetStringFromIDLine(cavebotIDselected, i)
     Next i
  End If
  Exit Sub
goterr:
 LogOnFile "errors.txt", "Error at UpdateValues(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub
Public Sub LoadCavebotChars()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim i As Long
  Dim firstC As Long
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
  cavebotIDselected = firstC
  UpdateValues
  Exit Sub
goterr:
 LogOnFile "errors.txt", "Error at LoadCavebotChars(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub

Public Sub AddScriptLine(str As String)
  Dim i As Long
  Dim startM As Long
  Dim endM As Long
  Dim initM As Long
  Dim lCaseStr As String
  startM = lstScript.ListIndex
  initM = lstScript.ListIndex
  lCaseStr = Trim$(LCase(str))
  
  ' setRetryAttacks, dropLootOnGround
  
  If Left$(lCaseStr, 15) = "setRetryAttacks" Then
    lblWarning.Caption = "WARNING - Some detection risk: " & vbCrLf & "setRetryAttacks"
  End If
  If Left$(lCaseStr, 16) = "dropLootOnGround" Then
    lblWarning.Caption = "WARNING - Some detection risk: " & vbCrLf & "dropLootOnGround"
  End If
  If Left$(lCaseStr, 14) = "fastexiva sell" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva sell"
  End If
  If Left$(lCaseStr, 21) = "saymessage exiva sell" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva sell"
  End If
  If Left$(lCaseStr, 21) = "sayintrade exiva sell" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva sell"
  End If
  If Left$(lCaseStr, 13) = "fastexiva buy" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva buy"
  End If
  If Left$(lCaseStr, 20) = "saymessage exiva buy" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva buy"
  End If
  If Left$(lCaseStr, 20) = "sayintrade exiva buy" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva buy"
  End If
  If Left$(lCaseStr, 11) = "fastexiva >" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva >"
  End If
  If Left$(lCaseStr, 18) = "saymessage exiva >" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva >"
  End If
  If Left$(lCaseStr, 18) = "sayintrade exiva >" Then
    lblWarning.Caption = "WARNING - High detection risk: " & vbCrLf & "exiva >"
  End If
  If startM < 0 Then
    AddIDLine cavebotIDselected, cavebotLenght(cavebotIDselected), str
    cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) + 1
  Else
    endM = cavebotLenght(cavebotIDselected) + 1
    startM = startM + 1
    For i = endM To startM Step -1
      AddIDLine cavebotIDselected, i, lstScript.List(i - 1)
    Next i
    AddIDLine cavebotIDselected, startM - 1, str
    cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) + 1
  End If
  UpdateValues
  lblEdit.Caption = "Edit current line ()"
  txtEdit.Text = ""

End Sub


Public Sub TurnCavebotState(idConnection As Integer, thisValue As Boolean)
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim shouldValue As Boolean
  Dim aRes As Long
  shouldValue = thisValue
  If (idConnection > 0) Then
    SelfDefenseID(idConnection) = 0
    If GameConnected(idConnection) = True Then
      DelayAttacks(idConnection) = 0
      GetProcessAllProcessIDs ' get new relations of process IDs
      If (thisValue = True) Then
        If (ProcessID(idConnection) = -1) Then
          ' need memory access to client
          lblInfo.Caption = "ERROR: CAN'T GET CLIENT PID!"
          OverwriteOnFile "debugpid.txt", debugPIDs(idConnection)
          shouldValue = False
        ElseIf (idConnection = cavebotIDselected) Then
          lblInfo.Caption = "running on pID:" & ProcessID(idConnection)
        End If
      End If
    Else
      shouldValue = False
    End If
    If (idConnection = cavebotIDselected) Then
      If (shouldValue = True) Then
        If (chkEnabled.Value <> 1) Then
          avoidC = True
          chkEnabled.Value = 1
          avoidC = False
        End If
      Else
        If (chkEnabled.Value <> 0) Then
          avoidC = True
          chkEnabled.Value = 0
          avoidC = False
        End If
      End If
    End If
    If (shouldValue = True) Then
      SpellKillHPlimit(idConnection) = 0
      SpellKillMaxHPlimit(idConnection) = 100
      LootAll(idConnection) = False
      PKwarnings(idConnection) = True
      OldLootMode(idConnection) = True
      ClientExecutingLongCommand(idConnection) = False
      AllowRepositionAtStart(idConnection) = 1
      AllowRepositionAtTrap(idConnection) = 1
      CavebotChaoticMode(idConnection) = 0
      exeLine(idConnection) = 0
      cavebotOnTrapGiveAlarm(idConnection) = False
      lastAttackedIDstatus(idConnection) = 0

      cancelAllMove(idConnection) = 0
      prevAttackState(idConnection) = False
      TurnsWithRedSquareZero(idConnection) = 0
      DangerPK(idConnection) = False
      DangerGM(idConnection) = False
      nextForcedDepotDeployRetry(idConnection) = 0
      somethingChangedInBps(idConnection) = False
      onDepotPhase(idConnection) = 0
      depotX(idConnection) = 0
      depotY(idConnection) = 0
      depotZ(idConnection) = 0
      doneDepotChestOpen(idConnection) = False
      depotTileID(idConnection) = 0
      depotS(idConnection) = 0
      lastDepotBPID(idConnection) = 0
      lastFloorTrap(idConnection) = -1
      lastDestX(idConnection) = 0
      lastDestY(idConnection) = 0
      lastDestZ(idConnection) = 0

        ignoreNext(idConnection) = -1 ' will reposition first
 
      '  ignoreNext(idConnection) = 0 ' force to start where script start
  
      RemoveAllMelee idConnection
      RemoveAllHMM idConnection
      RemoveAllSETUSEITEM idConnection
      RemoveAllExorivis idConnection
      RemoveAllAvoid idConnection
      RemoveAllShotType idConnection
      friendlyMode(idConnection) = 0
      requestLootBp(idConnection) = &HFF
      RemoveAllGoodLoot idConnection
      fishCounter(idConnection) = 0
      autoLoot(idConnection) = True
      cavebotOnDanger(idConnection) = -1
      cavebotOnGMclose(idConnection) = False
      'cavebotOnGMpause(idConnection) = False
      cavebotOnPLAYERpause(idConnection) = False
      CheatsPaused(idConnection) = False
      lastAttackedID(idConnection) = 0
      CavebotTimeWithSameTarget(idConnection) = GetTickCount()
      CavebotTimeStart(idConnection) = GetTickCount()
      RemoveAllIgnoredcreature idConnection
      maxAttackTime(idConnection) = 40000
      ChaotizeNextMaxAttackTime idConnection
      maxHit(idConnection) = 1000
      previousAttackedID(idConnection) = 0
      lastX(idConnection) = myX(idConnection)
      lastY(idConnection) = myY(idConnection)
      lastZ(idConnection) = myZ(idConnection)
      setFollowTarget(idConnection) = True
      waitCounter(idConnection) = GetTickCount()
      lblInfo.Caption = "running on pID:" & ProcessID(idConnection)
      RemoveAllClientSpamOrders idConnection
      pauseStacking(idConnection) = 0
      cavebotEnabled(idConnection) = True
      EnableMaxAttackTime(idConnection) = False
      AvoidReAttacks(idConnection) = True
      CavebotHaveSpecials(idConnection) = False
      CavebotLastSpecialMove(idConnection) = 0
      RemoveAllKillPriorities idConnection
      RemoveAllSpellKills idConnection
      cavebotOnGMpause(idConnection) = True ' new default since tibia 8.11
      ResetLooter idConnection
      'SendLogSystemMessageToClient idConnection, "Cavebot script started!"
      cavebotCurrentTargetPriority(idConnection) = 0
      usingPriorities(idConnection) = False
      DoEvents
    Else
      SpellKillHPlimit(idConnection) = 0
      SpellKillMaxHPlimit(idConnection) = 100
      TurnsWithRedSquareZero(idConnection) = 0
      LootAll(idConnection) = False
      PKwarnings(idConnection) = True
      OldLootMode(idConnection) = True
      ClientExecutingLongCommand(idConnection) = False
      CavebotChaoticMode(idConnection) = 0
      AvoidReAttacks(idConnection) = True
      cavebotOnTrapGiveAlarm(idConnection) = False
      cavebotEnabled(idConnection) = False
      EnableMaxAttackTime(idConnection) = False
      autoLoot(idConnection) = False
      exeLine(idConnection) = 0
      lastAttackedID(idConnection) = 0
      CavebotTimeWithSameTarget(idConnection) = GetTickCount()
      CavebotTimeStart(idConnection) = GetTickCount()
      RemoveAllIgnoredcreature idConnection
      maxAttackTime(idConnection) = 40000
      ChaotizeNextMaxAttackTime idConnection
      maxHit(idConnection) = 1000
      previousAttackedID(idConnection) = 0
      cavebotOnDanger(idConnection) = -1
      cavebotOnGMclose(idConnection) = False
      cavebotOnGMpause(idConnection) = False
      cavebotOnPLAYERpause(idConnection) = False
      RemoveAllMelee idConnection
      RemoveAllHMM idConnection
      RemoveAllSETUSEITEM idConnection
      RemoveAllAvoid idConnection
      RemoveAllShotType idConnection
      RemoveAllExorivis idConnection
      CavebotHaveSpecials(idConnection) = False
      CavebotLastSpecialMove(idConnection) = 0
      RemoveAllKillPriorities idConnection
      RemoveAllSpellKills idConnection
      usingPriorities(idConnection) = False
      cavebotCurrentTargetPriority(idConnection) = 0
      ResetLooter idConnection
      If (GameConnected(idConnection) = True) Then
        'SendLogSystemMessageToClient idConnection, "Cavebot script finished!"
        'DoEvents
      End If
    End If
  End If
  If (publicDebugMode = True) Then
    If (GameConnected(idConnection) = True) Then
      If (shouldValue = True) Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Cavebot was turned ON")
        DoEvents
      Else
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Cavebot was turned OFF")
        DoEvents
      End If
    End If
  End If
  Exit Sub
goterr:
  If idConnection > 0 Then
    frmMain.DoCloseActions (idConnection)
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Connection lose during TurnCavebotState on ID " & idConnection & " - CLOSING IT!"
  Else
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unknown error during TurnCavebotState : " & Err.Description
  End If
End Sub



Private Sub chkEnabled_Click()
  Dim beforeClickV As Boolean
  Dim idConnection As Integer
  Dim aRes As Long
  Dim startLine As Long
  Dim endLine As Long
  Dim lineID As Long
  Dim i As Long
    
  If cavebotIDselected > 0 Then
  If (avoidC = False) Then
    If chkEnabled.Value = 1 Then
    'DoUnifiedClickMove idconnection, (myX(cavebotIDselected) + randomNumberBetween(-10, 10)), (myY(cavebotIDselected) + randomNumberBetween(-10, 10)), myZ(cavebotIDselected)
    aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Cavebot ON$", cavebotIDselected, True)
    aRes = SendCustomSystemMessageToClient(cavebotIDselected, "Cavebot ON", &HB)
    'AddScriptLine "gotoScriptLine " & "0"
    TurnCavebotState cavebotIDselected, True
    Else
    aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Cavebot OFF$", cavebotIDselected, True)
    aRes = SendCustomSystemMessageToClient(cavebotIDselected, "Cavebot OFF", &HB)
      'cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) - 1
      'RemoveIDLine cavebotIDselected, cavebotLenght(cavebotIDselected)
    TurnCavebotState cavebotIDselected, False
    End If
  End If
End If
End Sub



Private Sub chkRunt_Click()
Dim idConnection As Integer
Dim aRes As String

'For idConnection = 1 To MAXCLIENTS
    If chkRunt.Value = 1 Then
        'chkEnabled.Value = 1
            'TimerRunt.enabled = True
            'SendLogSystemMessageToClient idConnection, "Cavebot script started!"
            aRes = SendCustomSystemMessageToClient(cavebotIDselected, "Cavebot script started", &HB)
    End If
    If chkRunt.Value = 0 Then
        'chkEnabled.Value = 0
            'TimerRunt.enabled = False
            'TurnCavebotState cavebotIDselected, False
            'SendLogSystemMessageToClient idConnection, "Cavebot script finished!"
            aRes = SendCustomSystemMessageToClient(cavebotIDselected, "Cavebot script finished", &HB)
    End If
'Next idConnection
End Sub

Private Sub cmbCharacter_Click()
  cavebotIDselected = cmbCharacter.ListIndex
  UpdateValues
End Sub













Private Sub cmdChange_Click()
    On Error GoTo goterr
    Dim lng1 As Long
    Dim lng2 As Long
    lng1 = CLng(txtMs.Text)
    lng2 = CLng(txtMs2.Text)
    If lng2 < lng1 Then
        GoTo goterr
    End If
    If lng1 < 20 Then
        GoTo goterr
    End If
    
    CavebotRECAST = lng1
    CavebotRECAST2 = lng2
    Me.Caption = "Cavebot - New timer = From " & CStr(lng1) & " to " & CStr(lng2) & " ms"
    Exit Sub
goterr:
    txtMs.Text = CStr(CavebotRECAST)
    txtMs2.Text = CStr(CavebotRECAST2)
    Me.Caption = "Cavebot - Wrong timer values!"
End Sub

Private Sub cmdChangeTimer_Click()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  MAX_LOCKWAIT = CLng(txtBlockSec)
  Exit Sub
goterr:
  MAX_LOCKWAIT = 30000
  txtBlockSec.Text = "30000"
End Sub

Private Sub cmdClose_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "closeConnection"
  End If
End Sub

Private Sub cmdComment_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "#" & CStr(txtComment.Text)
  End If
End Sub

Private Sub cmdDeleteSelected_Click()
  Dim startLine As Long
  Dim endLine As Long
  Dim i As Long
  If cavebotIDselected > 0 Then
    startLine = lstScript.ListIndex
    If startLine >= 0 Then
      endLine = cavebotLenght(cavebotIDselected) - 2
      For i = startLine To endLine
        AddIDLine cavebotIDselected, i, lstScript.List(i + 1)
      Next i
      cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) - 1
      RemoveIDLine cavebotIDselected, cavebotLenght(cavebotIDselected)
    End If
  End If
  UpdateValues
  lblEdit.Caption = "Edit current line ()"
  txtEdit.Text = ""
End Sub

Private Sub cmdDontRetryAttacks_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setDontRetryAttacks"
  End If
End Sub



Private Sub cmdDropLootOnGround_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "dropLootOnGround " & myX(cavebotIDselected) & "," & myY(cavebotIDselected) & "," & myZ(cavebotIDselected)
 End If
End Sub



Private Sub cmdFish_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "fishX " & CStr(txtFishTimes.Text)
  End If
End Sub



Private Sub cmdGotoScriptLine_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "gotoScriptLine " & CStr(txtGotoScriptLine.Text)
  End If
End Sub

Private Sub cmdIdlist_Click()
  frmIdlist.WindowState = vbNormal
  frmIdlist.Show
  frmIdlist.SetFocus
  SetWindowPos frmIdlist.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdIfOne_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "IfEnoughItemsGoto " & CStr(txtIfOne_Item.Text) & "," & _
     CStr(txtIfOne_Ammount.Text) & "," & CStr(txtIfOne_Goto.Text)
  End If
End Sub

Private Sub cmdIfTwo_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "IfFewItemsGoto " & CStr(txtIfTwo_Item.Text) & "," & _
     CStr(txtIfTwo_Ammount.Text) & "," & CStr(txtIfTwo_Goto.Text)
  End If
End Sub

Private Sub cmdLabel_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine ":" & CStr(txtLabel.Text)
  End If
End Sub

Private Sub cmdLoadCopyPaste_Click()
    On Error GoTo goterr
    Dim i As Long
    Dim ai As Long
    Dim pieces() As String
    Dim strLine As String
    Dim strtext As String
    Dim seguir As Boolean
    If cavebotIDselected > 0 Then
        lblInfo.Caption = "Waiting for copy/paste..."
        ClosedBoard = False
        frmBigText.lblText = "Copy the full script. Then paste it here." & vbCrLf & _
        "Finally press OK"
        frmBigText.Show
        SetWindowPos frmBigText.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        DisableBoardButtons
        While ClosedBoard = False
            DoEvents
        Wend
        EnableBoardButtons
        If CanceledBoard = False Then
            cavebotScript(cavebotIDselected).RemoveAll
            cavebotLenght(cavebotIDselected) = 0
            strtext = "" & frmBigText.txtBoard.Text
            If strtext <> "" Then
                pieces = Split(strtext, vbCrLf)
                i = 0
                For ai = 0 To UBound(pieces)
                  strLine = pieces(ai)
                  seguir = True
                  While seguir = True
                   If Len(strLine) < 1 Then
                     seguir = False
                   ElseIf Right$(strLine, 1) <> " " Then
                     seguir = False
                   Else
                     strLine = Left(strLine, (Len(strLine) - 1))
                   End If
                  Wend
                  If Len(strLine) >= 1 Then
                    AddIDLine cavebotIDselected, i, strLine
                    i = i + 1
                  End If
                Next ai
                cavebotLenght(cavebotIDselected) = i
                UpdateValues
                
                lblInfo.Caption = "Load done"
            End If
        Else
            lblInfo.Caption = ""
        End If
    Else
        lblInfo.Caption = "SELECT A CHARACTER FIRST!"
    End If
    Exit Sub
goterr:
    lblInfo.Caption = "Load failed, error " & CStr(Err.Number)
End Sub

Private Sub cmdLoadScript_Click()
  #If FinalMode Then
    On Error GoTo gotFerr
  #End If
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim filename As String
  Dim i As Long
  Dim seguir As Boolean
  Dim sp As Boolean
  #If FinalMode Then
    On Error GoTo goterr
  #End If
  lblWarning.Caption = ""
  Set fso = New scripting.FileSystemObject
  If cavebotIDselected > 0 Then
    cavebotScript(cavebotIDselected).RemoveAll
    cavebotLenght(cavebotIDselected) = 0
    filename = App.path & "\" & txtFile.Text
    If fso.FileExists(filename) = True Then
    
      fn = FreeFile
      Open filename For Input As #fn
      i = 0
      sp = False
      If EOF(fn) Then
        lblInfo.Caption = "File found, but it was empty!"
        sp = True
      Else
      While Not EOF(fn)
        Line Input #fn, strLine
        seguir = True
        While seguir = True
         If Len(strLine) < 1 Then
           seguir = False
         ElseIf Right$(strLine, 1) <> " " Then
           seguir = False
         Else
           strLine = Left(strLine, (Len(strLine) - 1))
         End If
        Wend
        If Len(strLine) >= 1 Then
          AddIDLine cavebotIDselected, i, strLine
          i = i + 1
        End If
      Wend
      End If
      Close #fn
      cavebotLenght(cavebotIDselected) = i
      If sp = False Then
      lblInfo.Caption = "Load OK"
      End If
    Else
      cavebotLenght(cavebotIDselected) = 0
      lblInfo.Caption = "Load failed - New file loaded"
    End If
  Else
    lblInfo.Caption = "SELECT A CHARACTER FIRST!"
  End If
  UpdateValues
  lblEdit.Caption = "Edit current line ()"
  txtEdit.Text = ""
  Exit Sub
goterr:
  lblInfo.Caption = "Load ERROR (" & Err.Number & "):" & Err.Description
  Exit Sub
gotFerr:
  lblInfo.Caption = "BIG load ERROR (" & Err.Number & "):" & Err.Description
  LogOnFile "errors.txt", "Error while loading a script: " & vbCrLf & _
  "Dim fso As Scripting.FileSystemObject <- This line failed with error number " & CStr(Err.Number) & " and error description: " & Err.Description
End Sub



Private Sub cmdMove_Click()
  AddCavebotMove
End Sub



Private Sub cmdMoveN_Click()
    AddCavebotMoveAE
End Sub

Private Sub cmdOnDangerGoto_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "onDangerGoto " & CStr(txtOnDangerGoto.Text)
  End If
End Sub

Private Sub cmdOnGMcloseConnection_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "onGMcloseConnection"
  End If
End Sub

Private Sub cmdOnGMpause_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "onGMpause"
  End If
End Sub

Private Sub cmdOnPlayerPause_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "onPLAYERpause-"
  End If
End Sub

Private Sub cmdOnTrapGiveAlarm_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "onTrapGiveAlarm"
  End If
End Sub

Private Sub cmdPutLootOnDepot_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "putLootOnDepot"
  End If
End Sub



Private Sub cmdReload_Click()
  ReloadFiles
End Sub

Private Sub cmdResetKillables_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "resetKill"
  End If
End Sub

Private Sub cmdResetLoot_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "resetLoot"
  End If
End Sub

Private Sub cmdRetryAttacks_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setRetryAttacks"
  End If
End Sub

Private Sub cmdSaveScript_Click()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim fn As Integer
  Dim limI As Long
  Dim i As Long
  If cavebotIDselected > 0 Then
    limI = cavebotLenght(cavebotIDselected) - 1
    fn = FreeFile
    Open App.path & "\" & txtFile.Text For Output As #fn
    For i = 0 To limI
      Print #fn, GetStringFromIDLine(cavebotIDselected, i)
    Next i
    Close #fn
    lblInfo.Caption = "Save OK"
  End If
  Exit Sub
goterr:
  lblInfo.Caption = "Save ERROR (" & Err.Number & "):" & Err.Description
End Sub

Private Sub cmdSayInTrade_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "sayInTrade " & CStr(txtSayMessage.Text)
  End If
End Sub

Private Sub cmdSetAny_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setAny"
  End If
End Sub

Private Sub cmdSetAvoidFront_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setAvoidFront " & CStr(txtAvoid.Text)
  End If
End Sub

Private Sub cmdSetBot_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setBot " & cmbSetOperator.Text & "=" & CStr(txtSetBotValue.Text)
  End If
End Sub

Private Sub cmdSetChaoticMovesOFF_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setChaoticMovesOFF"
  End If
End Sub

Private Sub cmdSetChaoticMovesON_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setChaoticMovesON"
  End If
End Sub

Private Sub cmdSetExoriMort_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setExoriMort " & CStr(txtMort.Text)
  End If
End Sub

Private Sub cmdSetExoriVis_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setExoriVis " & CStr(txtExori.Text)
  End If
End Sub

Private Sub cmdSetFollow_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setFollow"
  End If
End Sub

Private Sub cmdSetFriendly_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setFriendly"
  End If
End Sub

Private Sub cmdSetHmmKill_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setHmmKill " & CStr(txtSetHmmKill.Text)
  End If
End Sub

Private Sub cmdSetLoot_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setLoot " & CStr(txtSetLoot.Text)
  End If
End Sub

Private Sub cmdSetLootDistance_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setLootDistance " & CStr(txtSetLootDistance.Text)
  End If
End Sub

Private Sub cmdSetLootOff_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setLootOff"
  End If
End Sub

Private Sub cmdSetLootOn_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setLootOn"
  End If
End Sub

Private Sub cmdSetMaxAttackTimeMs_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "SetMaxAttackTimeMs " & CStr(txtSetMaxAttackTimeMs.Text)
  End If
End Sub

Private Sub cmdSetMaxHit_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "SetMaxHit " & CStr(txtSetMaxHit.Text)
  End If
End Sub

Private Sub cmdSetMeleeKill_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setMeleeKill " & CStr(txtSetMeleeKill.Text)
  End If
End Sub



Private Sub cmdSetNoFollow_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setNoFollow"
  End If
End Sub

Private Sub cmdSetPriority_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setPriority " & CStr(txtPriority1.Text) & ":" & CStr(txtPriority2.Text)
  End If
End Sub

Private Sub cmdSetSDkill_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setSDKill " & CStr(txtSetSDKill.Text)
  End If
End Sub

Private Sub cmdSetSpellKill_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "setSpellKill " & CStr(txtSetSpellKill_Creature.Text) & "," & CStr(txtSetSpellKill_Spell.Text) & "," & CStr(txtSetSpellKill_Dist.Text)
  End If
End Sub

Private Sub cmdSetVery_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "setVeryFriendly"
  End If
End Sub

Private Sub cmdStackItems_Click()
 If cavebotIDselected > 0 Then
    AddScriptLine "stackItems"
  End If
End Sub

Private Sub cmdUseItem_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "useItem " & myX(cavebotIDselected) & "," & myY(cavebotIDselected) & "," & myZ(cavebotIDselected)
  End If
End Sub



Private Sub cmdWait_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "waitX " & CStr(txtWait.Text)
  End If
End Sub

Private Sub cmdSayMessage_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "sayMessage " & CStr(txtSayMessage.Text)
  End If
End Sub

Public Sub ReloadFiles()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim fs As scripting.FileSystemObject
  Dim f As scripting.Folder
  Dim f1 As scripting.File
  Set fs = New scripting.FileSystemObject
  Set f = fs.GetFolder(App.path)
  txtFile.Clear
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "txt" Then
      If f1.name <> "code.txt" Then
        txtFile.AddItem f1.name
      End If
    End If
  Next
  txtFile.Text = "default.txt"
  Exit Sub
goterr:
  Me.Caption = "ERROR WITH FILESYSTEM OBJECT"
End Sub



Private Sub Command1_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "IfTrue (" & CStr(txtThing1.Text) & cmbOperator.Text & CStr(txtThing2.Text) & ") Goto " & Me.txtLineIFTRUE.Text
  End If
End Sub











Private Sub fastExiva_Click()
  If cavebotIDselected > 0 Then
    AddScriptLine "fastExiva " & CStr(txtFastExivaMessage.Text)
  End If
End Sub

Private Sub Form_Load()
 On Error GoTo goterr
    Me.txtMs.Text = CStr(CavebotRECAST)
    Me.txtMs2.Text = CStr(CavebotRECAST2)
    LoadCavebotChars
 With cmbOperator
 .Clear
 .AddItem "#number=#"
 .AddItem "#number<=#"
 .AddItem "#number>=#"
 .AddItem "#number<>#"
 .AddItem "#number<#"
 .AddItem "#number>#"
 .AddItem "#string=#"
 .AddItem "#string<>#"
 .Text = "#number<=#"
 End With
 
 With cmbSetOperator
 .Clear
 .AddItem "LootAll"
 .AddItem "PKwarnings"
 .AddItem "EnableMaxAttackTime"
 .AddItem "SpellKillHPlimit"
 .AddItem "SpellKillMaxHPlimit"
' .AddItem "OldLootMode"
' .AddItem "MINDELAYTOLOOT"
' .AddItem "MAXTIMEINLOOTQUEUE"
' .AddItem "MAXTIMETOREACHCORPSE"
 .AddItem "AllowRepositionAtStart"
 .AddItem "AllowRepositionAtTrap"
 .AddItem "AutoEatFood"
 .Text = "LootAll"
 End With
 Exit Sub
goterr:
  LogOnFile "errors.txt", "Could not load cavebot module. Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub











Private Sub lstScript_Click()
If cavebotIDselected > 0 Then
  If lstScript.ListIndex >= 0 Then
  lblEdit.Caption = "Edit current line (" & lstScript.ListIndex & ")"
    txtEdit.Text = lstScript.List(lstScript.ListIndex)
  End If
Else
  lblEdit.Caption = "Edit current line ()"
End If
End Sub




Public Sub scrollExorivis_Change()
  lblExorivisValue.Caption = CStr(scrollExorivis.Value) & " %"
End Sub

Public Sub scrollPkHeal_Change()
  lblPKhealValue.Caption = CStr(scrollPkHeal.Value) & " %"
End Sub





Private Sub TimerScript_Timer()
  Dim Sid As Integer
  Dim aRes As Long
  Dim gtc As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  gtc = GetTickCount()
  For Sid = 1 To HighestConnectionID
    If gtc > LastCavebotTime(Sid) Then
        LastCavebotTime(Sid) = gtc + randomNumberBetween(CavebotRECAST, CavebotRECAST2)
    
    If GameConnected(Sid) = True Then
    If (GotPacketWarning(Sid) = False) And (sentWelcome(Sid) = True) Then
    If (ClientExecutingLongCommand(Sid) = True) Then
      ' wait until long command is completed
      DoEvents
    ElseIf (GotKillOrder(Sid) = True) Then
      aRes = ThinkTheKill(Sid)
      DoEvents
    ElseIf (cavebotEnabled(Sid) = True) Then
      
      If (executingCavebot(Sid) = False) Then
        executingCavebot(Sid) = True
        ' end of script?
        If (exeLine(Sid) >= cavebotLenght(Sid)) Then
          ' finish and disable
          exeLine(Sid) = 0
          TurnCavebotState Sid, False
          executingCavebot(Sid) = False
          Exit Sub
        End If
        ' process line
        If Sid = 0 Then
          LogOnFile "errors.txt", "Error: value 1 to N returned 0!"
        End If
        ProcessScriptLine Sid
        executingCavebot(Sid) = False
      Else
        aRes = SendLogSystemMessageToClient(Sid, "Your CPU is overloaded. Skipping 1 turn of cavebot ...")
        DoEvents
      End If
    End If
    End If
    End If
    End If
  Next Sid
  Exit Sub
goterr:
  If Sid > 0 Then
    frmMain.DoCloseActions cavebotIDselected
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Connection lose during TimerScript_Timer() on ID " & CStr(Sid) & " - CLOSING IT!"
  Else
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unknown error during TimerScript_Timer() : " & Err.Description
  End If
End Sub



Private Sub txtBlockSec_Validate(Cancel As Boolean)
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  MAX_LOCKWAIT = CLng(txtBlockSec)
  Exit Sub
goterr:
  MAX_LOCKWAIT = 30000
  txtBlockSec.Text = "30000"
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    txtEdit_Validate False
  End If
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
If cavebotIDselected > 0 Then
  If lstScript.ListIndex >= 0 Then
    lstScript.List(lstScript.ListIndex) = txtEdit.Text
    ' update internal memory
    AddIDLine cavebotIDselected, lstScript.ListIndex, txtEdit.Text
  End If
End If
End Sub




Private Sub txtMs_Validate(Cancel As Boolean)
 Dim lngValue
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lngValue = CLng(txtMs.Text)
  If lngValue >= 10 And lngValue <= 500000 Then
    TimerScript.Interval = lngValue
  Else
    txtMs.Text = "300"
    TimerScript.Interval = 300
  End If
  Exit Sub
gotError:
  txtMs.Text = "300"
  TimerScript.Interval = 300
End Sub




