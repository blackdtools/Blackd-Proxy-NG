VERSION 5.00
Begin VB.Form frmTrainer 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exploring"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3300
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmTrainer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerC 
      Interval        =   300
      Left            =   2880
      Top             =   840
   End
   Begin VB.TextBox desired 
      Height          =   285
      Left            =   1620
      TabIndex        =   93
      Text            =   "exori frigo"
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Timer TimerDA 
      Interval        =   1200
      Left            =   2880
      Top             =   1260
   End
   Begin VB.CheckBox chkautoda 
      Caption         =   "Desired action"
      Height          =   255
      Left            =   180
      TabIndex        =   92
      Top             =   1380
      Width           =   1515
   End
   Begin VB.CheckBox chkautora 
      Caption         =   "Retry Attack"
      Height          =   195
      Left            =   180
      TabIndex        =   91
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Timer TimerEE 
      Interval        =   100
      Left            =   2880
      Top             =   420
   End
   Begin VB.CheckBox chkautoee 
      Caption         =   "Auto Exploring"
      Height          =   315
      Left            =   180
      TabIndex        =   90
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTrainer.frx":014A
      Left            =   7920
      List            =   "frmTrainer.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   88
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Items ID (hex)"
      Height          =   315
      Left            =   7920
      TabIndex        =   86
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtTrainerTimer2 
      Height          =   285
      Left            =   4680
      TabIndex        =   83
      Text            =   "1000"
      Top             =   7000
      Width           =   735
   End
   Begin VB.CommandButton cmdChangeTrainerTimer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CHANGE"
      Height          =   285
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Drop all loot of your containers on ground (house or guildhall)"
      Top             =   7020
      Width           =   855
   End
   Begin VB.TextBox txtTrainerTimer 
      Height          =   285
      Left            =   3240
      TabIndex        =   80
      Text            =   "300"
      Top             =   7000
      Width           =   735
   End
   Begin VB.Timer timerTrainer 
      Interval        =   300
      Left            =   8520
      Top             =   720
   End
   Begin VB.CommandButton cmdLastAttackedID 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Get ID of last attacked"
      Height          =   285
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Drop all loot of your containers on ground (house or guildhall)"
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtExceptionID 
      Height          =   285
      Left            =   8220
      TabIndex        =   72
      Text            =   "0"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CheckBox chkAvoidID 
      BackColor       =   &H00000000&
      Caption         =   "Avoid attacking the monster with this ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8220
      TabIndex        =   71
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CheckBox chkDance14min 
      BackColor       =   &H00000000&
      Caption         =   "Dance at 15 minutes autologout warning"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8220
      TabIndex        =   70
      Top             =   5760
      Width           =   3495
   End
   Begin VB.TextBox txtMinAllowedHP 
      Height          =   285
      Left            =   12480
      TabIndex        =   68
      Text            =   "50"
      Top             =   5445
      Width           =   375
   End
   Begin VB.CheckBox chkStopLowHp 
      BackColor       =   &H00000000&
      Caption         =   "Stop attacking target until regen"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8220
      TabIndex        =   67
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   10
      Left            =   12360
      TabIndex        =   65
      Text            =   "1"
      Top             =   4710
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   9
      Left            =   12360
      TabIndex        =   64
      Text            =   "1"
      Top             =   4350
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   8
      Left            =   12360
      TabIndex        =   63
      Text            =   "1"
      Top             =   3990
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   7
      Left            =   12360
      TabIndex        =   62
      Text            =   "1"
      Top             =   3630
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   6
      Left            =   12360
      TabIndex        =   61
      Text            =   "1"
      Top             =   3270
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   5
      Left            =   12360
      TabIndex        =   60
      Text            =   "1"
      Top             =   2910
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   4
      Left            =   12360
      TabIndex        =   59
      Text            =   "1"
      Top             =   2550
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   3
      Left            =   12360
      TabIndex        =   58
      Text            =   "1"
      Top             =   2190
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   57
      Text            =   "1"
      Top             =   1830
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   55
      Text            =   "1"
      Top             =   1470
      Width           =   735
   End
   Begin VB.TextBox txtSlotAmmount 
      Height          =   285
      Index           =   0
      Left            =   12840
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   10
      Left            =   7080
      TabIndex        =   53
      Text            =   "00 00"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   9
      Left            =   11400
      TabIndex        =   51
      Text            =   "CD 0C"
      Top             =   4350
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Ring:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   9600
      TabIndex        =   50
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   8
      Left            =   11400
      TabIndex        =   49
      Text            =   "CD 0C"
      Top             =   3990
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Boots:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   9600
      TabIndex        =   48
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   7
      Left            =   11400
      TabIndex        =   47
      Text            =   "CD 0C"
      Top             =   3630
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Legs:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   9600
      TabIndex        =   46
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   6
      Left            =   11400
      TabIndex        =   45
      Text            =   "CD 0C"
      Top             =   3270
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Left hand:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   9600
      TabIndex        =   44
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   5
      Left            =   11400
      TabIndex        =   43
      Text            =   "CD 0C"
      Top             =   2910
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Right hand:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   9600
      TabIndex        =   42
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   4
      Left            =   11400
      TabIndex        =   41
      Text            =   "CD 0C"
      Top             =   2550
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Chest:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   40
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   3
      Left            =   11400
      TabIndex        =   39
      Text            =   "CD 0C"
      Top             =   2190
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Backpack:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   38
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   2
      Left            =   11400
      TabIndex        =   37
      Text            =   "CD 0C"
      Top             =   1830
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Neck:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   36
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   1
      Left            =   11400
      TabIndex        =   35
      Text            =   "CD 0C"
      Top             =   1470
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Head:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   34
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtSlotRefill 
      Height          =   285
      Index           =   0
      Left            =   12840
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H00000000&
      Caption         =   "Error:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   12840
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtMaxPickUp 
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Text            =   "99999"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtPickupID 
      Height          =   285
      Left            =   7080
      TabIndex        =   25
      Text            =   "00 00"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   8
      Left            =   6750
      Picture         =   "frmTrainer.frx":014E
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   7
      Left            =   5715
      Picture         =   "frmTrainer.frx":3520
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   6
      Left            =   4680
      Picture         =   "frmTrainer.frx":68F2
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   5
      Left            =   6750
      Picture         =   "frmTrainer.frx":9CC4
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4515
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   4
      Left            =   5715
      Picture         =   "frmTrainer.frx":D096
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4515
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   3
      Left            =   4680
      Picture         =   "frmTrainer.frx":10468
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4515
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   2
      Left            =   6750
      Picture         =   "frmTrainer.frx":1383A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   1
      Left            =   5715
      Picture         =   "frmTrainer.frx":16C0C
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1040
   End
   Begin VB.CommandButton cmdPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   0
      Left            =   4680
      Picture         =   "frmTrainer.frx":19FDE
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   8
      Left            =   6750
      Picture         =   "frmTrainer.frx":1D3B0
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   7
      Left            =   5715
      Picture         =   "frmTrainer.frx":20782
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   6
      Left            =   4680
      Picture         =   "frmTrainer.frx":23B54
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   5
      Left            =   6750
      Picture         =   "frmTrainer.frx":26F26
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   4
      Left            =   5715
      Picture         =   "frmTrainer.frx":2A2F8
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   3
      Left            =   4680
      Picture         =   "frmTrainer.frx":2D6CA
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   2
      Left            =   6750
      Picture         =   "frmTrainer.frx":30A9C
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   1
      Left            =   5715
      Picture         =   "frmTrainer.frx":33E6E
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.CommandButton cmdNoPickup 
      BackColor       =   &H00000000&
      Height          =   1040
      Index           =   0
      Left            =   4680
      Picture         =   "frmTrainer.frx":37240
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1040
   End
   Begin VB.Frame frmPickDestination 
      BackColor       =   &H00000000&
      Caption         =   "Destination of items"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   3135
      Begin VB.OptionButton OptionDest 
         BackColor       =   &H00000000&
         Caption         =   "Ammo slot"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   78
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptionDest 
         BackColor       =   &H00000000&
         Caption         =   "Any backpack"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptionDest 
         BackColor       =   &H00000000&
         Caption         =   "Right hand"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   76
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionDest 
         BackColor       =   &H00000000&
         Caption         =   "Left hand"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Text            =   "-"
      Top             =   180
      Width           =   1815
   End
   Begin VB.CheckBox chkSlotRefill 
      BackColor       =   &H80000018&
      Caption         =   "Refill Arrow slot"
      Height          =   255
      Index           =   10
      Left            =   6240
      TabIndex        =   52
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CheckBox chkEnableTrainer 
      BackColor       =   &H80000018&
      Caption         =   "Active collectitems"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   89
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000003&
      X1              =   60
      X2              =   3180
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000003&
      X1              =   60
      X2              =   3180
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   2
      X1              =   60
      X2              =   60
      Y1              =   1800
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   3180
      X2              =   3180
      Y1              =   1800
      Y2              =   660
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   9480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000018&
      Caption         =   "Arrow ID :"
      Height          =   255
      Left            =   6240
      TabIndex        =   87
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   9480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000D&
      X1              =   9480
      X2              =   9480
      Y1              =   1200
      Y2              =   2040
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   6120
      Y1              =   1200
      Y2              =   2040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   9480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      X1              =   9480
      X2              =   9480
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   9480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   6120
      X2              =   6120
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Label lblInEffect 
      BackColor       =   &H00000000&
      Caption         =   "Values in effect: from 300 ms to 1000 ms"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   85
      Top             =   7320
      Width           =   6255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "ms"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   84
      Top             =   7035
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "ms to"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   81
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Global trainer timer: Randomized from"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   79
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   7200
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   4440
      TabIndex        =   75
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "WARNING: it will lose that ID if it dies or 'relog' !"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8220
      TabIndex        =   74
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "% hp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12960
      TabIndex        =   69
      Top             =   5490
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "MISC TRAINER OPTIONS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   66
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "AUTOREFILL PLAYER SLOTS FROM BACKPACKS, when you have less than X items there:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9600
      TabIndex        =   29
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "X value"
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
      Left            =   12360
      TabIndex        =   56
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "player slot"
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
      Left            =   9960
      TabIndex        =   33
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   11400
      TabIndex        =   32
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Max items that you can carry:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   6645
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "(Default is spear ID)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   26
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Caption         =   "Item ID :"
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click on the picture. Leave a spear on the squares allowed for auto pick up."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000014&
      Caption         =   "Char :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblGlobalEvents 
      BackColor       =   &H00000000&
      Caption         =   "PICK UP SPEARS OR OTHER ITEMS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   2340
      Width           =   3135
   End
End
Attribute VB_Name = "frmTrainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Public Sub UpdateValues()
  Dim goodValues As Boolean
  Dim strTmp As String
  Dim i As Integer
  goodValues = False
  If trainerIDselected > 0 Then
    If GameConnected(trainerIDselected) = True Then
      goodValues = True
    End If
  End If
  If goodValues = True Then
    txtPickupID.Text = GoodHex(TrainerOptions(trainerIDselected).spearID_b1) & _
     " " & GoodHex(TrainerOptions(trainerIDselected).spearID_b2)
    If OptionDest(TrainerOptions(trainerIDselected).spearDest).Value = False Then
      OptionDest(TrainerOptions(trainerIDselected).spearDest).Value = True
    End If
    txtMaxPickUp.Text = "9999"
        txtPickupID.Text = txtPickupID.Text
    If OptionDest(2).Value = False Then
      OptionDest(2).Value = True
    End If
    If OptionDest(0).Value = True Then
      OptionDest(0).Value = False
    End If
    For i = 0 To 8
      If (TrainerOptions(trainerIDselected).AllowedSides(i) = True) Then
        cmdPickup(i).Visible = True
        cmdNoPickup(i).Visible = False
      Else
        cmdPickup(i).Visible = False
        cmdNoPickup(i).Visible = True
      End If
    Next i
    For i = 1 To 10
      If chkSlotRefill(i).Value <> TrainerOptions(trainerIDselected).PlayerSlots(i).cheked Then
        chkSlotRefill(i) = TrainerOptions(trainerIDselected).PlayerSlots(i).cheked
      End If
      strTmp = GoodHex(TrainerOptions(trainerIDselected).PlayerSlots(i).itemID_b1) & _
       " " & GoodHex(TrainerOptions(trainerIDselected).PlayerSlots(i).itemID_b2)
      txtSlotRefill(i).Text = strTmp
      txtSlotAmmount(i).Text = CStr(TrainerOptions(trainerIDselected).PlayerSlots(i).xvalue)
    Next i
    If chkStopLowHp.Value <> TrainerOptions(trainerIDselected).misc_stoplowhp Then
      chkStopLowHp.Value = TrainerOptions(trainerIDselected).misc_stoplowhp
    End If
    If chkDance14min.Value <> TrainerOptions(trainerIDselected).misc_dance_14min Then
      chkDance14min.Value = TrainerOptions(trainerIDselected).misc_dance_14min
    End If
    If chkAvoidID.Value <> TrainerOptions(trainerIDselected).misc_avoidID Then
      chkAvoidID.Value = TrainerOptions(trainerIDselected).misc_avoidID
    End If
    If chkEnableTrainer.Value <> TrainerOptions(trainerIDselected).enabled Then
      chkEnableTrainer.Value = TrainerOptions(trainerIDselected).enabled
    End If
    
    
    
    
    txtExceptionID.Text = CStr(TrainerOptions(trainerIDselected).idToAvoid)
    txtMinAllowedHP.Text = CStr(TrainerOptions(trainerIDselected).stoplowhpHP)
  Else 'defaults
    txtPickupID.Text = "D7 0B"
    If OptionDest(2).Value = False Then
      OptionDest(2).Value = True
    End If
    If OptionDest(0).Value = True Then
      OptionDest(0).Value = False
    End If
    txtMaxPickUp.Text = "99999"
    For i = 0 To 8
      cmdPickup(i).Visible = False
      cmdNoPickup(i).Visible = True
    Next i
    For i = 1 To 10
      If chkSlotRefill(i).Value <> 0 Then
        chkSlotRefill(i).Value = 0
      End If
      txtSlotRefill(i).Text = "00 00"
      txtSlotAmmount(i).Text = "1"
    Next i
    If chkStopLowHp.Value <> 0 Then
      chkStopLowHp.Value = 0
    End If
    If chkDance14min.Value <> 0 Then
      chkDance14min.Value = 0
    End If
    If chkAvoidID.Value <> 0 Then
      chkAvoidID.Value = 0
    End If
    If chkEnableTrainer.Value <> 0 Then
      chkEnableTrainer.Value = 0
    End If
    txtExceptionID.Text = "0"
    txtMinAllowedHP.Text = "50"
  End If
End Sub

Public Sub LoadTrainerChars()
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
      cmbCharacter.AddItem "-", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  trainerIDselected = firstC
  UpdateValues
End Sub

Private Sub chkautoda_Click()
If lock_chkautoda = False Then
If trainerIDselected > 0 Then
  If chkautoda.Value = 1 Then
    RuneMakerOptions(trainerIDselected).autoda = True
  Else
    RuneMakerOptions(trainerIDselected).autoda = False
  End If
End If
End If
End Sub

Private Sub chkautoee_Click()
If lock_chkautoee = False Then
If trainerIDselected > 0 Then
  If chkautoee.Value = 1 Then
    RuneMakerOptions(trainerIDselected).autoee = True
    'RuneMakerOptions(trainerIDselected).autoarme6 = True
  Else
    RuneMakerOptions(trainerIDselected).autoee = False
    'RuneMakerOptions(trainerIDselected).autoarme6 = False
  End If
End If
End If

  'If (trainerIDselected > 0) Then
  '  RuneMakerOptions(trainerIDselected).autoee = chkautoee.Value
  'End If

End Sub

Private Sub chkautora_Click()
If lock_chkautora = False Then
If trainerIDselected > 0 Then
  If chkautora.Value = 1 Then
    RuneMakerOptions(trainerIDselected).autora = True
  Else
    RuneMakerOptions(trainerIDselected).autora = False
  End If
End If
End If
End Sub

Private Sub chkAvoidID_Click()
  If (trainerIDselected > 0) Then
    TrainerOptions(trainerIDselected).misc_avoidID = chkAvoidID.Value
  End If
End Sub

Private Sub chkDance14min_Click()
  If (trainerIDselected > 0) Then
    TrainerOptions(trainerIDselected).misc_dance_14min = chkDance14min.Value
  End If
End Sub

Private Sub chkEnableTrainer_Click()

  If (runemakerIDselected > 0) Then
    TrainerOptions(runemakerIDselected).enabled = chkEnableTrainer.Value
  End If
End Sub

Private Sub chkSlotRefill_Click(index As Integer)
  Dim thenewvalue As Long
  thenewvalue = chkSlotRefill(index).Value
  If ((runemakerIDselected > 0) And (index > 0)) Then
    TrainerOptions(runemakerIDselected).PlayerSlots(index).cheked = thenewvalue
  End If
End Sub

Private Sub chkStopLowHp_Click()
  If (trainerIDselected > 0) Then
    TrainerOptions(trainerIDselected).misc_stoplowhp = chkStopLowHp.Value
  End If
End Sub

Private Sub cmbCharacter_Click()
  trainerIDselected = cmbCharacter.ListIndex
   If trainerIDselected = 0 Then
      trainerIDselected = 1
  End If
  'UpdateValues
  If trainerIDselected > 0 Then
      UpdateValues
  End If
  
    If RuneMakerOptions(trainerIDselected).autoee = True Then
      chkautoee.Value = 1
    Else
      chkautoee.Value = 0
    End If
    
    If RuneMakerOptions(trainerIDselected).autora = True Then
      chkautora.Value = 1
    Else
      chkautora.Value = 0
    End If
    
    If RuneMakerOptions(trainerIDselected).autoda = True Then
      chkautoda.Value = 1
    Else
      chkautoda.Value = 0
    End If
  
End Sub

Private Sub cmdChangeTrainerTimer_Click()
    On Error GoTo ignoretheerror
    Dim lngNewOne As Long
    Dim lngNewOne2 As Long
    lngNewOne = CLng(txtTrainerTimer.Text)
    lngNewOne2 = CLng(txtTrainerTimer2.Text)
    If lngNewOne < 10 Then
        'timerTrainer.Interval = lngNewOne
        GoTo ignoretheerror
    End If
    If lngNewOne > lngNewOne2 Then
        GoTo ignoretheerror
    End If
    TrainerTimer1 = lngNewOne
    TrainerTimer2 = lngNewOne2
    lblInEffect.Caption = "Values in effect: from " & _
     CStr(TrainerTimer1) & " ms to " & CStr(TrainerTimer2) & " ms"
    Exit Sub
ignoretheerror:
    txtTrainerTimer.Text = CStr(TrainerTimer1)
    txtTrainerTimer2.Text = CStr(TrainerTimer2)
    lblInEffect.Caption = "Error in input fields. Values in effect: from " & _
     CStr(TrainerTimer1) & " ms to " & CStr(TrainerTimer2) & " ms"
End Sub

Private Sub cmdLastAttackedID_Click()
  If trainerIDselected > 0 Then
    If GameConnected(trainerIDselected) = True Then
      txtExceptionID.Text = CStr(currTargetID(trainerIDselected))
    End If
  End If
End Sub

Private Sub cmdNoPickup_Click(index As Integer)
  cmdNoPickup(index).Visible = False
  cmdPickup(index).Visible = True
  If (trainerIDselected > 0) Then
    TrainerOptions(trainerIDselected).AllowedSides(index) = True
  End If
  cmdPickup(index).SetFocus
End Sub

Private Sub cmdPickup_Click(index As Integer)
  cmdPickup(index).Visible = False
  cmdNoPickup(index).Visible = True
  If (trainerIDselected > 0) Then
    TrainerOptions(trainerIDselected).AllowedSides(index) = False
  End If
  cmdNoPickup(index).SetFocus
End Sub


Private Sub Combo1_Click()
Dim index As Integer
'Dim Index As Integer
'Dim i As Integer
'ammo.ListIndex = ammo.NewIndex
If Combo1.ListIndex = 0 Then
'desired.Text = "77 0D" 'arrow
txtSlotRefill(10).Text = "77 0D"
ElseIf Combo1.ListIndex = 1 Then
'desired.Text = "79 0D" 'burst
txtSlotRefill(10).Text = "79 0D"
ElseIf Combo1.ListIndex = 2 Then
'desired.Text = "C4 1C" 'sniper
txtSlotRefill(10).Text = "C4 1C"
ElseIf Combo1.ListIndex = 3 Then
'desired.Text = "AB 37" 'tarsal
txtSlotRefill(10).Text = "AB 37"
ElseIf Combo1.ListIndex = 4 Then
'desired.Text = "C5 1C" 'onyx
txtSlotRefill(10).Text = "C5 1C"
ElseIf Combo1.ListIndex = 5 Then
'desired.Text = "0F 3F" 'envenon
txtSlotRefill(10).Text = "0F 3F"
ElseIf Combo1.ListIndex = 6 Then
'desired.Text = "B1 3D" 'crystal
txtSlotRefill(10).Text = "B1 3D"
ElseIf Combo1.ListIndex = 7 Then
'desired.Text = "76 0D" ' bolt
txtSlotRefill(10).Text = "76 0D"
ElseIf Combo1.ListIndex = 8 Then
'desired.Text = "C3 1C" 'piercing
txtSlotRefill(10).Text = "C3 1C"
ElseIf Combo1.ListIndex = 9 Then
'desired.Text = "AC 37" 'vortex
txtSlotRefill(10).Text = "AC 37"
ElseIf Combo1.ListIndex = 10 Then
'desired.Text = "0E 3F" 'drill
txtSlotRefill(10).Text = "0E 3F"
ElseIf Combo1.ListIndex = 11 Then
'desired.Text = "7A 0D" 'pbolt
txtSlotRefill(10).Text = "7A 0D"
ElseIf Combo1.ListIndex = 12 Then
'desired.Text = "0D 3F" 'prismatic
txtSlotRefill(10).Text = "0D 3F"
ElseIf Combo1.ListIndex = 13 Then
'desired.Text = "80 19" 'infernal
txtSlotRefill(10).Text = "80 19"
End If

End Sub

Private Sub Command1_Click()
  frmIdlist.WindowState = vbNormal
  frmIdlist.Show
  frmIdlist.SetFocus
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
  trainerIDselected = 0
  Me.txtTrainerTimer = CStr(TrainerTimer1)
  Me.txtTrainerTimer2 = CStr(TrainerTimer2)
  Me.timerTrainer.Interval = 300
  lblInEffect.Caption = "Values in effect: from " & _
     CStr(TrainerTimer1) & " ms to " & CStr(TrainerTimer2) & " ms"
  UpdateValues
  chkSlotRefill(10).Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub






Private Sub OptionDest_Click(index As Integer)
  If ((trainerIDselected > 0) And (index > 0) And OptionDest(index).Value = True) Then
    TrainerOptions(trainerIDselected).spearDest = index
  End If
End Sub







Private Sub TimerC_Timer()
Dim idConnection As Integer
Dim i As Integer
Dim Value As Integer
Dim target As String
'Dim aRes As Long

'target = "test"

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       
  '    If (GotKillOrder(idConnection) = True) And (RuneMakerOptions(idConnection).autoee = True) Then
  '        RuneMakerOptions(idConnection).autoarme4 = True
  '    Else
  '        'RuneMakerOptions(idConnection).autoarme4 = False
  '    End If
  '
  '    If (GotKillOrder(idConnection) = False) And (RuneMakerOptions(idConnection).autoee = True) Then
  '        RuneMakerOptions(idConnection).autoarme4 = False
  '    Else
  '        'RuneMakerOptions(idConnection).autoarme4 = False
  '    End If
      
      If (RuneMakerOptions(idConnection).autora = True) Then 'And (RuneMakerOptions(idConnection).autoarme4 = True) Then
           'ProcessKillOrder idConnection, target
            'ProcessKillOrder3 (idConnection)
      End If
            
    End If
    Next idConnection
End Sub

Private Sub TimerDA_Timer()
Dim idConnection As Integer
Dim i As Integer
Dim Value As Integer
Dim aRes As Long
Dim magia As String

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       
       magia = desired.Text
      If (RuneMakerOptions(idConnection).autoda = True) And (RuneMakerOptions(idConnection).autora = True) And (RuneMakerOptions(idConnection).autoarme4 = True) And (RuneMakerOptions(idConnection).autoarme5 = True) Then
           aRes = ExecuteInTibia(magia, idConnection, True)
      Else 'nothing
      End If
      
    End If
    Next idConnection
End Sub

Private Sub TimerEE_Timer()
Dim idConnection As Integer
Dim i As Integer
Dim val1 As Long
Dim val2 As Long
Dim val3 As Long
Dim val4 As Long
Dim val5 As Long

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
      If (RuneMakerOptions(idConnection).autoee = True) And (RuneMakerOptions(idConnection).autoarme4 = False) Then

                 val4 = randomNumberBetween(1, 5) * 5
                 val5 = randomNumberBetween(1, 5) * 5
                 val1 = CLng((myX(idConnection) + val4) - val5)
                 val2 = CLng((myY(idConnection) + val4) - val5)
                 val3 = CLng(myZ(idConnection))
                 PerformMove idConnection, val1, val2, val3
                
      End If
    End If
    Next idConnection
End Sub



Private Sub timerTrainer_Timer()
  #If FinalMode Then
  On Error GoTo fatalError
  #End If
  Dim idConnection As Integer
  Dim slotID As Byte
  Dim tileID As Long
  Dim gotTileID As Long
  Dim iRes As Integer
  Dim sPos As Byte
  Dim sfoundhere As Byte
  Dim aRes As Long
  Dim sCheat As String
  Dim xpos As Long
  Dim ypos As Long
  Dim cPacket() As Byte
  Dim am As Byte
  Dim res1 As TypeSearchItemResult2
  Dim carring As Long
  Dim wanted As Long
  Dim blockingItem As Byte
  If frmHardcoreCheats.chkApplyCheats.Value = 0 Then
    Exit Sub ' fixed since 10.3
  End If
  Me.timerTrainer.Interval = 300
  For idConnection = 1 To MAXCLIENTS
  If (TrainerOptions(idConnection).enabled = 1) Then
  'If (TrainerOptions(idConnection).enabled = True) Then
  
    If ((GameConnected(idConnection) = True) And (sentWelcome(idConnection) = True) And (GotPacketWarning(idConnection) = False)) Then
  
        
      tileID = GetTheLong(TrainerOptions(idConnection).spearID_b1, _
       TrainerOptions(idConnection).spearID_b2)
      If ((TrainerOptions(idConnection).spearID_b1 = 0) And (TrainerOptions(idConnection).spearID_b2 = 0)) Then
        wanted = &H64
      Else
         carring = SearchAmmount(idConnection, TrainerOptions(idConnection).spearID_b1, _
          TrainerOptions(idConnection).spearID_b2)
          wanted = 9999 '- carring
          
      End If
      If wanted > 0 Then
      For slotID = 0 To 8
        If TrainerOptions(idConnection).AllowedSides(slotID) = False Then 'aqui col
          sfoundhere = &HFF
          If TibiaVersionLong >= 780 Then
          'new pickup
          
          
          For sPos = 1 To 10
            
           gotTileID = GetTheLong(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t1, _
              Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t2)
           If (Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).dblID = 0) Then ' not person
           
           
           If ((TrainerOptions(idConnection).spearID_b1 = 0) And (TrainerOptions(idConnection).spearID_b2 = 0)) Then
             If DatTiles(gotTileID).pickupable = True Then
               sfoundhere = sPos
               'Exit For
             End If
             If (TibiaVersionLong > 860) Then
               If gotTileID > 0 Then
                If (DatTiles(gotTileID).notMoveable = False) And (DatTiles(gotTileID).alwaysOnTop = False) And (DatTiles(gotTileID).groundtile = False) Then
                  sfoundhere = sPos
                  'Exit For
                End If
               End If
             End If
           ElseIf (Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t1 = TrainerOptions(idConnection).spearID_b1) And _
             (Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t2 = TrainerOptions(idConnection).spearID_b2) Then
             sfoundhere = sPos
             'Exit For
           End If
           If ((DatTiles(gotTileID).moreAlwaysOnTop = False) And (DatTiles(gotTileID).alwaysOnTop = False)) Then
           Exit For 'exit in any case now, unless it was a person
           End If
           End If
          Next sPos
           
           
           
          Else 'old pickup
          For sPos = 1 To 10
            gotTileID = GetTheLong(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t1, _
              Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t2)
           If ((TrainerOptions(idConnection).spearID_b1 = 0) And (TrainerOptions(idConnection).spearID_b2 = 0)) Then
             If DatTiles(gotTileID).pickupable = True Then
               sfoundhere = sPos
               Exit For
             End If
           ElseIf (Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t1 = TrainerOptions(idConnection).spearID_b1) And _
             (Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sPos).t2 = TrainerOptions(idConnection).spearID_b2) Then
             sfoundhere = sPos
             Exit For
           End If
          Next sPos
          End If
          If sfoundhere < &HFF Then
            'aRes = SendLogSystemMessageToClient(idconnection, "Found one at " & CStr(slotID) & " spos " & GoodHex(sfoundhere))
            'DoEvents
            xpos = myX(idConnection) - 1 + (slotID Mod 3)
            ypos = myY(idConnection) - 1 + (slotID \ 3)
            If DatTiles(gotTileID).haveExtraByte Then
              am = CByte(MinV(CLng(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t3), wanted))
            Else
              am = &H1
            End If
            Select Case TrainerOptions(idConnection).spearDest
            Case 0 'left
              If publicDebugMode = True Then
                aRes = SendLogSystemMessageToClient(idConnection, "Autopickup is moving " & str(CLng(am)) & " to left hand")
                DoEvents
              End If
              sCheat = "78 " & GoodHex(LowByteOfLong(xpos)) & " " & GoodHex(HighByteOfLong(xpos)) & " " & _
               GoodHex(LowByteOfLong(ypos)) & " " & GoodHex(HighByteOfLong(ypos)) & " " & _
               GoodHex(CByte(myZ(idConnection))) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t1) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t2) & _
               " " & GoodHex(sfoundhere) & " FF FF 06 00 00 " & GoodHex(am)
              'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " >> " & sCheat
              SafeCastCheatString "timer_trainer1", idConnection, sCheat
            Case 1 'right
              If publicDebugMode = True Then
                aRes = SendLogSystemMessageToClient(idConnection, "Autopickup is moving " & str(CLng(am)) & " to right hand")
                DoEvents
              End If
              sCheat = "78 " & GoodHex(LowByteOfLong(xpos)) & " " & GoodHex(HighByteOfLong(xpos)) & " " & _
               GoodHex(LowByteOfLong(ypos)) & " " & GoodHex(HighByteOfLong(ypos)) & " " & _
               GoodHex(CByte(myZ(idConnection))) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t1) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t2) & _
               " " & GoodHex(sfoundhere) & " FF FF 05 00 00 " & GoodHex(am)
              'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " >> " & sCheat
              SafeCastCheatString "timer_trainer2", idConnection, sCheat
            Case 2 'any bp
              If publicDebugMode = True Then
                aRes = SendLogSystemMessageToClient(idConnection, "Autopickup is moving " & str(CLng(am)) & " to any bp")
                DoEvents
              End If
              res1 = SearchItemDestinationForLoot(idConnection, Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t1, _
               Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t2, &HFF)
              If res1.foundcount > 0 Then
              sCheat = "78 " & GoodHex(LowByteOfLong(xpos)) & " " & GoodHex(HighByteOfLong(xpos)) & " " & _
               GoodHex(LowByteOfLong(ypos)) & " " & GoodHex(HighByteOfLong(ypos)) & " " & _
               GoodHex(CByte(myZ(idConnection))) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t1) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t2) & _
               " " & GoodHex(sfoundhere) & " FF FF " & GoodHex(&H40 + res1.bpID) & " 00 " & GoodHex(res1.slotID) & " " & GoodHex(am)
              'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " >> " & sCheat
              SafeCastCheatString "timer_trainer3", idConnection, sCheat
              End If
            Case 3 'ammo
              If publicDebugMode = True Then
                aRes = SendLogSystemMessageToClient(idConnection, "Autopickup is moving " & str(CLng(am)) & " to ammo")
                DoEvents
              End If
              sCheat = "8 " & GoodHex(LowByteOfLong(xpos)) & " " & GoodHex(HighByteOfLong(xpos)) & " " & _
               GoodHex(LowByteOfLong(ypos)) & " " & GoodHex(HighByteOfLong(ypos)) & " " & _
               GoodHex(CByte(myZ(idConnection))) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t1) & " " & _
               GoodHex(Matrix(-1 + (slotID \ 3), -1 + (slotID Mod 3), myZ(idConnection), idConnection).s(sfoundhere).t2) & _
               " " & GoodHex(sfoundhere) & " FF FF 0A 00 00 " & GoodHex(am)
              'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " >> " & sCheat
              SafeCastCheatString "timer_trainer4", idConnection, sCheat
            End Select
            Exit For
          End If
        End If
      Next slotID
     ' End If
     '   For slotID = 1 To EQUIPMENT_SLOTS
     '       If TrainerOptions(idConnection).PlayerSlots(slotID).cheked = 1 Then
     '       tileID = GetTheLong(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1, _
     '       TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2)
     '           gotTileID = GetTheLong(mySlot(idConnection, slotID).t1, mySlot(idConnection, slotID).t2)
     '           If tileID <> 0 Then
     '               If tileID <= highestDatTile Then
     '                   If DatTiles(tileID).stackable = True Then
     '                       If (tileID <> gotTileID) Or _
     '                       ((tileID = gotTileID) And _
     '                       (CLng(mySlot(idConnection, slotID).t3) < TrainerOptions(idConnection).PlayerSlots(slotID).xvalue)) Then
     '                           iRes = ExecuteInTibia("exiva #" & _
     '                           GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1) & " " & _
     '                           GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2) & " " & _
     '                           GoodHex(slotID), idConnection, True)
     '                       End If
     '                   Else ' not stackable : only replace if you have nothing there. since b8.79
     '                       If (gotTileID = 0) Then
     '                           iRes = ExecuteInTibia("exiva #" & _
     '                           GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1) & " " & _
     '                           GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2) & " " & _
     '                           GoodHex(slotID), idConnection, True)
     '                       End If
     '                   End If
     '               End If
     '           End If
            End If
     '   Next slotID
    End If
  End If
  'aq teste
  For slotID = 1 To EQUIPMENT_SLOTS
            If TrainerOptions(idConnection).PlayerSlots(slotID).cheked = 1 Then
            tileID = GetTheLong(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1, _
            TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2)
                gotTileID = GetTheLong(mySlot(idConnection, slotID).t1, mySlot(idConnection, slotID).t2)
                If tileID <> 0 Then
                    If tileID <= highestDatTile Then
                        If DatTiles(tileID).stackable = True Then
                            If (tileID <> gotTileID) Or _
                            ((tileID = gotTileID) And _
                            (CLng(mySlot(idConnection, slotID).t3) < TrainerOptions(idConnection).PlayerSlots(slotID).xvalue)) Then
                                iRes = ExecuteInTibia("exiva #" & _
                                GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1) & " " & _
                                GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2) & " " & _
                                GoodHex(slotID), idConnection, True)
                            End If
                        Else ' not stackable : only replace if you have nothing there. since b8.79
                            If (gotTileID = 0) Then
                                iRes = ExecuteInTibia("exiva #" & _
                                GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b1) & " " & _
                                GoodHex(TrainerOptions(idConnection).PlayerSlots(slotID).itemID_b2) & " " & _
                                GoodHex(slotID), idConnection, True)
                            End If
                        End If
                    End If
                End If
            End If
        Next slotID
  Next idConnection
  Exit Sub
fatalError:
  LogOnFile "errors.txt", "Fatal error caught at timeTrainer. Code " & CStr(Err.Number) & " : " & Err.Description
End Sub

Private Sub txtExceptionID_Change()
  Dim res As Double
  If (trainerIDselected > 0) Then
    res = safeConvertStringToDouble(txtExceptionID.Text)
    TrainerOptions(trainerIDselected).idToAvoid = res
  End If
End Sub

Private Sub txtMaxPickUp_Change()
  Dim res As Long
  If (trainerIDselected > 0) Then
    res = 9999999
    TrainerOptions(trainerIDselected).maxitems = res
  End If
End Sub

Private Sub txtMinAllowedHP_Change()
  Dim res As Long
  If (trainerIDselected > 0) Then
    res = safeConvertStringToLong(txtMinAllowedHP.Text)
    TrainerOptions(trainerIDselected).stoplowhpHP = res
  End If
End Sub

Private Sub txtPickupID_Change()
  Dim res As TypePairOfBytes
  Dim strTmp As String
  
  'strTmp = frmRunemaker.txtPickupID.Text
    
  If runemakerIDselected > 0 Then
   ' res = safeConvertStringToPairOfBytes(strTmp)
   ' TrainerOptions(runemakerIDselected).spearID_b1 = res.b1
   ' TrainerOptions(runemakerIDselected).spearID_b2 = res.b2
  End If
End Sub



Private Sub txtSlotAmmount_Change(index As Integer)
  Dim res As Long
  If ((trainerIDselected > 0) And (index > 0)) Then
    res = safeConvertStringToLong(txtSlotAmmount(index).Text)
    TrainerOptions(trainerIDselected).PlayerSlots(index).xvalue = res
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

Private Sub txtTrainerTimer_Change()
    On Error GoTo ignoretheerror
    Dim lngNewOne As Long
    lngNewOne = CLng(txtTrainerTimer.Text)
    If lngNewOne >= 10 Then
        timerTrainer.Interval = lngNewOne
    End If
    Exit Sub
ignoretheerror:
End Sub


Private Sub txtTrainerTimer2_Change()
    On Error GoTo ignoretheerror
    Dim lngNewOne As Long
    lngNewOne = CLng(txtTrainerTimer2.Text)
    If lngNewOne >= 10 Then
        timerTrainer.Interval = lngNewOne
    End If
    Exit Sub
ignoretheerror:
End Sub


