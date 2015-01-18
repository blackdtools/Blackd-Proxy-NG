VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blackd NG"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin JwldButn2b.JeweledButton JeweledButton24 
      Height          =   300
      Left            =   5400
      TabIndex        =   55
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "4"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton JeweledButton25 
      Height          =   300
      Left            =   5760
      TabIndex        =   54
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "5"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton JeweledButton23 
      Height          =   300
      Left            =   5040
      TabIndex        =   53
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "3"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton JeweledButton22 
      Height          =   300
      Left            =   4680
      TabIndex        =   52
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "2"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton JeweledButton21 
      Height          =   300
      Left            =   4320
      TabIndex        =   51
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "1"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdMCtools 
      Height          =   300
      Left            =   0
      TabIndex        =   50
      Top             =   900
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "MC tools"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdUnknownFeature 
      Height          =   300
      Left            =   3240
      TabIndex        =   49
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Conditions"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdStealth 
      Height          =   300
      Left            =   1080
      TabIndex        =   48
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Aimbot"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdBroadcast 
      Height          =   300
      Left            =   1080
      TabIndex        =   46
      Top             =   600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Broadcast"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdAdvanced 
      Height          =   300
      Left            =   5400
      TabIndex        =   45
      Top             =   600
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   529
      Caption         =   "Config"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdNews 
      Height          =   300
      Left            =   5400
      TabIndex        =   44
      Top             =   300
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   529
      Caption         =   "News"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdLoad 
      Height          =   300
      Left            =   4320
      TabIndex        =   43
      Top             =   600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Load"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdSav 
      Height          =   300
      Left            =   4320
      TabIndex        =   42
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Save"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdLaunchTibiaMC 
      Height          =   300
      Left            =   3240
      TabIndex        =   41
      Top             =   600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Tibia MC"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdStopAlarm 
      Height          =   300
      Left            =   3240
      TabIndex        =   40
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Stop Alarm"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdLogs 
      Height          =   300
      Left            =   2160
      TabIndex        =   39
      Top             =   600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Proxy"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdWarbot 
      Height          =   300
      Left            =   2160
      TabIndex        =   38
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Lists"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdTargeting 
      Height          =   300
      Left            =   2160
      TabIndex        =   37
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Exploring"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdHotkeys 
      Height          =   300
      Left            =   1080
      TabIndex        =   36
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Hotkeys"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdCavebot 
      Height          =   300
      Left            =   0
      TabIndex        =   35
      Top             =   600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Cavebot"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdRunemaker 
      Height          =   300
      Left            =   0
      TabIndex        =   34
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Extras"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdHardcoreCheats 
      Height          =   300
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Healing"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin VB.CommandButton cmdAdvanced2 
      BackColor       =   &H80000014&
      Caption         =   "Config"
      Height          =   333
      Left            =   5340
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":058A
      TabIndex        =   25
      Top             =   4620
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdNews2 
      BackColor       =   &H80000018&
      Caption         =   "Help"
      Height          =   333
      Left            =   5340
      MaskColor       =   &H80000012&
      Picture         =   "frmMenu.frx":15E2
      TabIndex        =   29
      Top             =   4260
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "Load"
      Height          =   333
      Left            =   5280
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSav2 
      Caption         =   "Save"
      Height          =   330
      Left            =   5220
      MaskColor       =   &H80000012&
      TabIndex        =   26
      Top             =   2220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLaunchTibiaMC2 
      BackColor       =   &H80000014&
      Caption         =   "Tibia MC"
      Height          =   333
      Left            =   2760
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":32D7
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStopAlarm2 
      BackColor       =   &H80000014&
      Caption         =   "Stop Alarm"
      Height          =   333
      Left            =   3720
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":42BE
      TabIndex        =   28
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogs2 
      BackColor       =   &H80000018&
      Caption         =   "Proxy"
      Height          =   333
      Left            =   4020
      MaskColor       =   &H80000012&
      Picture         =   "frmMenu.frx":52F6
      TabIndex        =   22
      Top             =   2340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdWarbot2 
      BackColor       =   &H80000014&
      Caption         =   "Lists"
      Height          =   333
      Left            =   5520
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":6037
      TabIndex        =   24
      Top             =   2700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdTarget 
      BackColor       =   &H80000018&
      Caption         =   "Targeting"
      Height          =   333
      Left            =   4260
      MaskColor       =   &H80000012&
      TabIndex        =   31
      Top             =   2700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRunemaker2 
      BackColor       =   &H8000000E&
      Caption         =   "Extras"
      Height          =   333
      Left            =   2700
      MaskColor       =   &H80000012&
      Picture         =   "frmMenu.frx":75D8
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdHotkeys2 
      BackColor       =   &H8000000E&
      Caption         =   "Hotkeys"
      Height          =   333
      Left            =   1440
      MaskColor       =   &H80000012&
      Picture         =   "frmMenu.frx":8903
      TabIndex        =   23
      Top             =   1740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdHardcoreCheats2 
      BackColor       =   &H80000018&
      Caption         =   "Healing"
      Height          =   333
      Left            =   2160
      MaskColor       =   &H80000012&
      Picture         =   "frmMenu.frx":993D
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdHPmana 
      Caption         =   "Heal?"
      Enabled         =   0   'False
      Height          =   315
      Left            =   360
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":ACB8
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdVIPsupport 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to VIP support page"
      Height          =   315
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdBroadcast2 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   3540
      Picture         =   "frmMenu.frx":BD25
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdLaunchTibia 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Picture         =   "frmMenu.frx":DA05
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdStealth2 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Picture         =   "frmMenu.frx":10660
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAd 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   2400
      Picture         =   "frmMenu.frx":11C90
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnknownFeature2 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   2100
      Picture         =   "frmMenu.frx":12A8D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdMagebomb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   4080
      Picture         =   "frmMenu.frx":13E07
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdEvents 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   1200
      Picture         =   "frmMenu.frx":14D68
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutorial 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      Picture         =   "frmMenu.frx":15BC1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheats 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   6480
      Picture         =   "frmMenu.frx":16AD1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdTrainer 
      BackColor       =   &H00000000&
      Caption         =   "Collect"
      Enabled         =   0   'False
      Height          =   333
      Left            =   900
      Picture         =   "frmMenu.frx":17BDE
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCavebot2 
      BackColor       =   &H80000014&
      Caption         =   "Cavebot"
      Height          =   333
      Left            =   780
      MaskColor       =   &H80000017&
      Picture         =   "frmMenu.frx":18ABC
      TabIndex        =   20
      Top             =   2100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin JwldButn2b.JeweledButton cmdDebugs 
      Height          =   300
      Left            =   1080
      TabIndex        =   47
      Top             =   900
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Debug tools"
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
      BorderColor_Hover=   16744576
      BorderColor_Inner=   16777215
   End
   Begin VB.Label Label4 
      Caption         =   "v.1.00"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   5520
      TabIndex        =   32
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "If you purchased us any gold in the last month, we give you VIP support"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   4320
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblAltSite 
      BackColor       =   &H00000000&
      Caption         =   "www.blackdtools.es"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblForum 
      BackColor       =   &H00000000&
      Caption         =   "[forum]"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUpdates 
      BackColor       =   &H00000000&
      Caption         =   "[updates]"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3300
      TabIndex        =   5
      Top             =   6540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMainSite 
      BackColor       =   &H00000000&
      Caption         =   "www.blackdtools.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3420
      TabIndex        =   4
      Top             =   5940
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Official sites:"
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
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Blackd Proxy!  Remember to download updates in our web from time to time."
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
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4980
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Close"
      End
      Begin VB.Menu mPopShowTibia 
         Caption         =   "&Show Tibia"
      End
      Begin VB.Menu mPopHideTibia 
         Caption         =   "&Hide Tibia"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
#Const BlockCavebot = 0
#Const BlockTools = 0
#Const BlockRunemaker = 0
#Const BlockAllCheats = 0
#Const KeepRemote = 0
#Const DoSave = 1
Private AutoloadUsable As Boolean
Private Const CteAutoloadSubfolder As String = "autoload"
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Option Explicit

Private Sub cmdAd_Click()
'Dim a

  'a = ShellExecute(Me.hWnd, "Open", "https://blackdtools.com/worldtrade.php", &O0, &O0, SW_NORMAL)

End Sub

Private Sub cmdAdvanced_Click()
  ' show Advanced form
  frmAdvanced.WindowState = vbNormal
  frmAdvanced.Show
  frmAdvanced.SetFocus
  SetWindowPos frmAdvanced.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdBroadcast_Click()
  frmBroadcast.WindowState = vbNormal
  frmBroadcast.Show
  frmBroadcast.SetFocus
  SetWindowPos frmBroadcast.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub



'#Const FinalMode =0
'#Const BlockCavebot = 1
'#Const BlockTools = 1
'#Const BlockRunemaker = 1
'#Const BlockAllCheats = 1
'#Const KeepRemote = 1
'#Const DoSave = 0
Private Sub cmdCavebot_Click()
  ' show cavebot form
  frmCavebot.WindowState = vbNormal
  frmCavebot.Show
  frmCavebot.SetFocus
  SetWindowPos frmCavebot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdCheats_Click()
  ' show Tools form
'  frmCheats.WindowState = vbNormal
'  frmCheats.Show
'  frmCheats.SetFocus
End Sub

Private Sub cmdDebugs_Click()
  ' This menu is only displayed in debug mode (#Const FinalMode = 0)
  ' It is only usefull for programmers.
  frmCheats.WindowState = vbNormal
  frmCheats.Show
  frmCheats.SetFocus
  SetWindowPos frmCheats.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdEvents_Click()
'  frmEvents.WindowState = vbNormal
'  frmEvents.Show
'  frmEvents.SetFocus
End Sub



Private Sub cmdHardcoreCheats_Click()
  ' show Cheats form
  frmHardcoreCheats.WindowState = vbNormal
  frmHardcoreCheats.Show
  frmHardcoreCheats.SetFocus
  frmHardcoreCheats.UpdateValues
  SetWindowPos frmHardcoreCheats.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub





Private Sub cmdHotkeys_Click()
  ' show hotkeys
  frmHotkeys.WindowState = vbNormal
  frmHotkeys.Show
  frmHotkeys.SetFocus
  SetWindowPos frmHotkeys.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdHPmana_Click()
  ' show Advanced form
  'frmHPmana.WindowState = vbNormal
  'frmHPmana.Show
  'frmHPmana.SetFocus
End Sub




Private Sub cmdLaunchTibia_Click()
  ' open Shoot Fruits in default web navigator
'  Dim X
'  X = ShellExecute(Me.hWnd, "Open", "http://shootfruits.com", &O0, &O0, SW_NORMAL)
  
'    Dim res As String
'    Dim tpath As String
'    tpath = TibiaExePath
'    If tpath = "" Then
'        Label3.Caption = "FILESYSTEM ERROR"
'        Exit Sub
'    End If
'    res = LaunchTibia(tpath, False)
'    If res <> "" Then
'        Label3.Caption = "TIBIA NOT FOUND"
'        Exit Sub
'    End If
End Sub

Private Sub cmdLaunchTibiaMC_Click()
    Dim res As String
    Dim tpath As String
    tpath = TibiaExePath
    If tpath = "" Then
        Label3.Caption = "FILESYSTEM ERROR"
        Exit Sub
    End If
    res = LaunchTibia(tpath, True)
    If res <> "" Then
        Label3.Caption = "TIBIA NOT FOUND"
        Exit Sub
    End If
End Sub

Private Sub cmdLoad_Click()
Dim aRes As Long
Dim louade As String
Dim idConnection As Integer
Dim i As Integer
    For i = 1 To MAXCLIENTS
        idConnection = i
        louade = "exiva load"
        aRes = ExecuteInTibia(louade, idConnection, True)
        aRes = ExecuteInTibia(louade, idConnection, True)
        DoEvents
        Next i
End Sub

Private Sub cmdLogs_Click()
  ' show main form
  frmMain.WindowState = vbNormal
  frmMain.Show
  frmMain.SetFocus
  SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdMagebomb_Click()
'  frmMagebomb.WindowState = vbNormal
'  frmMagebomb.Show
'  frmMagebomb.SetFocus
End Sub

Private Sub cmdMCtools_Click()
  frmMCtools.WindowState = vbNormal
  frmMCtools.Show
  frmMCtools.SetFocus
  SetWindowPos frmMCtools.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

'Private Sub cmdMagebot_Click()
'    Dim res As String
'    Dim tpath As String
'    Dim tfile As String
'    tpath = MagebotPath
'    tfile = MagebotExe
'    If tpath = "" Then
'        Label3.Caption = "FILESYSTEM ERROR"
'        Exit Sub
'    End If
'    res = LaunchFileNormalWay(tpath, tfile)
'    If res <> "" Then
'        Label3.Caption = "TIBIA NOT FOUND"
'        Exit Sub
'    End If
'End Sub

Private Sub cmdNews_Click()
  frmNews.WindowState = vbNormal
  frmNews.Show
  frmNews.SetFocus
  SetWindowPos frmNews.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub



Private Sub cmdRunemaker_Click()
  ' show Runemaker form
  frmRunemaker.WindowState = vbNormal
  frmRunemaker.Show
  frmRunemaker.SetFocus
  SetWindowPos frmRunemaker.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdSav_Click()
Dim aRes As Long
Dim salve As String
Dim idConnection As Integer
Dim i As Integer
'frmHardcoreCheats.UpdateValues

    For i = 1 To MAXCLIENTS
        idConnection = i
        salve = "exiva save"
        aRes = ExecuteInTibia(salve, idConnection, True)
        DoEvents
        Next i
End Sub

Private Sub cmdStealth_Click()
  frmStealth.WindowState = vbNormal
  frmStealth.Show
  frmStealth.SetFocus
  SetWindowPos frmStealth.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdStopAlarm_Click()
  ' stop alarms
  Dim mcid As Integer
  For mcid = 1 To MAXCLIENTS
    DangerPK(mcid) = False
    DangerGM(mcid) = False
    DangerPlayer(mcid) = False
    LogoutTimeGM(mcid) = 0
    moveRetry(mcid) = 0
    RemoveSpamOrder mcid, 1
    UHRetryCount(mcid) = 0
    logoutAllowed(mcid) = 0
  Next mcid
  ChangePlayTheDangerSound False
  PlayPMSound = False
  PlayMsgSound = False
End Sub

Private Sub cmdTarget_Click()

MsgBox "This is a Beta Version, many features are unavailable. To use Targeting use the Blackd Module on Cavebot button until new Version.", _
vbExclamation + vbOKOnly, _
currentAppName

  ' show target form
  'frmTarget.WindowState = vbNormal
  'frmTarget.Show
  'frmTarget.SetFocus
  'frmTarget.ReloadFiles

End Sub



Private Sub cmdTargeting_Click()
  frmTrainer.WindowState = vbNormal
  frmTrainer.Show
  frmTrainer.SetFocus
  SetWindowPos frmTrainer.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdTutorial_Click()
  ' open tutorial in default web navigator
'  Dim X
'  X = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.com/forum/showthread.php?t=221", &O0, &O0, SW_NORMAL)
End Sub




Private Sub cmdUnknownFeature_Click()
  frmCondEvents.WindowState = vbNormal
  frmCondEvents.Show
  frmCondEvents.SetFocus
  SetWindowPos frmCondEvents.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdVIPsupport_Click()
 ' open tutorial in default web navigator
'  Dim X
'  X = ShellExecute(Me.hWnd, "Open", "https://blackdtools.com/vip.php", &O0, &O0, SW_NORMAL)

End Sub

Private Sub cmdWarbot_Click()
  frmWarbot.WindowState = vbNormal
  frmWarbot.Show
  frmWarbot.SetFocus
  SetWindowPos frmWarbot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub





'Private Sub Command1_Click()
'  Dim tibiaclient As Long
'  Dim res As Long
'  tibiaclient = 0
'  Do
'        tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
'        If tibiaclient = 0 Then
'            Exit Do
'        Else
'            res = ReadCurrentAddress(tibiaclient, adrSelectedCharIndex, -1, True)
'            MsgBox ("SEL CHAR=" & CStr(res))
'        End If
'  Loop
'End Sub

Private Sub Form_Load()
  Dim pok As Boolean
  #If FinalMode = 1 Then
    Me.cmdDebugs.enabled = False
    Me.cmdDebugs.Visible = False
  #End If
  
  Me.Label4.Caption = "v " & ProxyVersion
  If thisShouldNotBeLoading = 0 Then
    Unload Me
    Exit Sub
  End If
  
  If TibiaVersionLong >= 841 Then
    frmMenu.cmdMagebomb.enabled = False
  End If
  
  'If IamAdmin = True Then
  '  lblAdminInfo.Caption = "Running as admin: " & App.EXEName & ".exe"
  'Else
  '  lblAdminInfo.Caption = "Running as user: " & App.EXEName & ".exe"
  'End If
  Me.Caption = frmMain.Caption
  frmMain.Caption = "Proxy (connection and logs)"

  CornerMessage = "If you purchased us any gold in the last month, we give you VIP support"
 
  Label3.Caption = CornerMessage
  Label3.ForeColor = CornerColor
  ApplyLimits
  Me.Show
  Me.Refresh
  With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = currentAppName & vbNullChar
  End With
  Shell_NotifyIcon NIM_ADD, nid
  DoEvents
  If FirstExecute = True Then
    cmdTutorial_Click
  End If
  pok = True
  If MyPriorityID <> 2 Then
    pok = UpdateMyPriority()
  End If
  If (TibiaPriorityID <> 2) And (pok = True) Then
    pok = UpdateTibiaPriority()
  End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload

  ' Unload all
'  If thisShouldNotBeLoading = 0 Then
'    Exit Sub
'  End If
'  If confirmedExit = False Then
'    Cancel = True
'    ToggleTopmost frmConfirm.hwnd, True
'    frmConfirm.Show
'    frmConfirm.SetFocus
'    Exit Sub
'  End If
'  #If DoSave Then
'    If LoadWasCompleted = True Then ' check to avoid ini corruption if there was an unexpected fail at loading
'      frmMain.WriteIni
'    End If
'  #End If
'  BlockUnload = 0
'  frmMain.Timer1.enabled = False 'should not be needed ... just in case
'  frmEvents.timerScheduledActions.enabled = False
'  'frmTrainer.timerTrainer.enabled = False
'  frmCondEvents.timerCheck.enabled = False
' 'this removes the icon from the system tray
'  Shell_NotifyIcon NIM_DELETE, nid
'  Refresh 'ensure icon is deleted from tray
'  'LogOnFile "debug.txt", "Ended by user"
'  Unload frmMain
'  Unload frmCheats
'  Unload frmBigText
'  Unload frmCavebot
'  Unload frmTrueMap
'  Unload frmBackpacks
'  Unload frmRunemaker
'  Unload frmHardcoreCheats
'  Unload frmAdvanced
'  Unload frmHotkeys
'  Unload frmMagebomb
'  Unload frmScreenshot
'  Unload frmCondEvents
'  Unload frmHPmana
'  Unload frmNews
'  Unload frmStealth
'  Unload frmBroadcast
'  Set DirectX = Nothing
'  Set DX = Nothing
'
'  End 'ends all subthreads of this process
End Sub







Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
         Single, Y As Single)
  'this procedure receives the callbacks from the System Tray icon.
  Dim result As Long
  Dim msg As Long
  'the value of X will vary depending upon the scalemode setting
  If Me.ScaleMode = vbPixels Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If
  
  Select Case msg
  Case WM_LBUTTONUP        '514 restore form window
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Hide
    Me.Show
  Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Show
  Case WM_RBUTTONUP        '517 display popup menu
    result = SetForegroundWindow(Me.hwnd)
    Me.PopupMenu Me.mPopupSys
  End Select
End Sub

Private Sub Form_Resize()
  ' this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then
    Me.Hide
  End If
End Sub



Private Sub JeweledButton21_Click()

If GameConnected(1) = True Then
Me.Caption = "Blackd NG - " & frmRunemaker.cmbCharacter.List(1)
End If

End Sub

Private Sub JeweledButton22_Click()

If GameConnected(2) = True Then
Me.Caption = "Blackd NG - " & frmRunemaker.cmbCharacter.List(2)
End If

End Sub

Private Sub JeweledButton23_Click()

If GameConnected(3) = True Then
Me.Caption = "Blackd NG - " & frmRunemaker.cmbCharacter.List(3)
End If

End Sub

Private Sub JeweledButton24_Click()

If GameConnected(4) = True Then
Me.Caption = "Blackd NG - " & frmRunemaker.cmbCharacter.List(4)
End If

End Sub

Private Sub JeweledButton25_Click()

If GameConnected(5) = True Then
Me.Caption = "Blackd NG - " & frmRunemaker.cmbCharacter.List(5)
End If

End Sub





Private Sub lblAltSite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim a
'If Button = 1 Then
' a = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.es/index.php", &O0, &O0, SW_NORMAL)
'End If
End Sub

Private Sub lblForum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a
If Button = 1 Then
  a = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/index.php", &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lblMainSite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim a
'If Button = 1 Then
' a = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.com/index.php", &O0, &O0, SW_NORMAL)
'End If
End Sub



Private Sub lblUpdates_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim a
'If Button = 1 Then
'  a = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.com/freedownloads.php", &O0, &O0, SW_NORMAL)
'End If
End Sub

Private Sub mPopExit_Click()
  ' exit by tray menu
  Dim btemp As Integer
  btemp = 0
  If confirmedExit = False Then
    frmConfirm.Show
    frmConfirm.SetFocus
    SetWindowPos frmConfirm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Exit Sub
  End If
  'called when user clicks the popup menu Exit command
End Sub

Private Sub mPopRestore_Click()
  'called when the user clicks the popup menu Restore command
  Dim result As Long
  Me.WindowState = vbNormal
  result = SetForegroundWindow(Me.hwnd)
  Me.Show
End Sub

Private Sub mPopShowTibia_Click()
  SetTibiaClientsVisible True
End Sub

Private Sub mPopHideTibia_Click()
  SetTibiaClientsVisible False
End Sub

Public Sub ApplyLimits()
' If compiling a limited version, then disable and hide some options
Dim save1 As Long
#If BlockRunemaker Then
With frmRunemaker
.UseRightHand.enabled = False
.UseLeftHand.enabled = False
.chkActivate.enabled = True
.chkFood.enabled = False
.chkManaFluid.enabled = False
.chkautoUtamo.enabled = False
.chkautoAp.enabled = False
.chkautossa.enabled = False
.chkautopmax.enabled = False
.chkautotar.enabled = False
.chkautoSdt.enabled = False
.chkautoDan.enabled = False
.chkautodd.enabled = False
.chkautoee.enabled = False
.chkautoarme4.enabled = False
.chkautoarme5.enabled = False
.chkautoarme6.enabled = False
.chkautora.enabled = False
.chkautoda.enabled = False
.chkautoxray.enabled = False
.chkautodk.enabled = False
.chkautogHur.enabled = False
.chkautoHur.enabled = False
.chkautoPM.enabled = False
.chkautoaim.enabled = False
.chkautoUE.enabled = False
.chklocktrigger.enabled = False
.chkLogoutDangerAny.enabled = False
.chkLogoutDangerCurrent.enabled = False
.chkLogoutOutRunes.enabled = False
.chkWaste.enabled = False
.chkssap.enabled = False
.chkerg.enabled = False
.chkmsgSound.enabled = False
.chkmsgSound2.enabled = False
.txtAction1.enabled = False
.txtManaAction1.enabled = False
.Text2.enabled = False
.Text3.enabled = False
.txtAction2.enabled = False
.txtManaAction2.enabled = False
.txtSoulAction2.enabled = False
.lstFriends.enabled = False
.cmdLoad.enabled = False
.cmdSave.enabled = False
.txtFile.enabled = False
.txtAddFriend.enabled = False
.cmdAddFriend.enabled = False
.cmdRemoveFriend.enabled = False
.ChkDangerSound.Value = 0
.ChkDangerSound.enabled = False
.chkCloseSound.Value = 0
.chkOnDangerSS.Value = 0
.chkCloseSound.enabled = False
.cmdStopAlarm.enabled = False
.cmdApply.enabled = False
.cmdDebug.enabled = False
End With
frmMenu.cmdRunemaker2.enabled = False
#End If
#If BlockRunemaker Then
With frmCavebot
.chkEnabled.enabled = False
.chkChangePkHeal.Value = 0
.chkChangePkHeal.enabled = False
End With
frmMenu.cmdCavebot2.enabled = False
#End If
#If BlockTools Then
frmCheats.chkInspectTileID.Value = 0
frmCheats.chkInspectTileID.enabled = False
#End If
save1 = frmHardcoreCheats.chkAcceptSDorder.Value
#If BlockAllCheats Then
With frmHardcoreCheats
.txtRemoteLeader.Text = LimitedLeader
.chkLogoutIfDanger.Value = 0
.chkLogoutIfDanger.enabled = False
.chkReveal.Value = 0
.chkReveal.enabled = False
.chkLight.Value = 0
.chkLight.enabled = False
.chkAutoHeal.Value = 0
.chkAutoHeal.enabled = False
.chkAutoVita.Value = 0
.chkAutoVita.enabled = False
.chkAcceptSDorder.Value = 0
.chkAcceptSDorder.enabled = False
.chkColorEffects.Value = 0
.chkColorEffects.enabled = False
.cmdOpenTrueRadar.enabled = False
.cmdUpdateMap.enabled = False
.cmdOpenBackpacks.enabled = False
.chkLogoutIfDanger.Visible = False
.chkReveal.Visible = False
.chkLight.Visible = False
.chkAutoHeal.Visible = False
.chkAutoVita.Visible = False
.chkAcceptSDorder.Visible = False
.chkColorEffects.Visible = False
.cmdOpenTrueRadar.Visible = False
.cmdUpdateMap.Visible = False
.cmdOpenBackpacks.Visible = False
.chkApplyCheats.Visible = False
.cmdReset.Visible = False
.Line3.Visible = False
.scrollLight.Visible = False
.lblLightValue.Visible = False
.scrollHP.Visible = False
.lblHPvalue.Visible = False
.scrollHP3.Visible = False
.lblHPvalue3.Visible = False
.scrollHP4.Visible = False
.lblHPvalue4.Visible = False
.scrollHP2.Visible = False
.lblHPvalue2.Visible = False
.txtOrder.Visible = False
.lblOrder2.Visible = False
.lblRead.Visible = False
.cmbOrderType.Visible = False
.lblOn.Visible = False
.lblLeader.Visible = False
.txtRemoteLeader.Visible = False
.txtCommands.Visible = False
.chkColorEffects.Visible = False
.cmdOpenTrueRadar.Visible = False
.cmdUpdateMap.Visible = False
.chkLockOnMyFloor.Visible = False
.chkOnTop.Visible = False
.cmdOpenBackpacks.Visible = False
.lblChar.Visible = False
.cmbCharacter.Visible = False
.lblYourPos.Visible = False
.lblPosition.Visible = False
.chkManualUpdate.Visible = False
.chkUpdateMs.Visible = False
.chkAutoUpdateMap.Visible = False
.Label1.Visible = False
.lblArraySelected.Visible = False
.cmdMs.Visible = False
.cmdChange.Visible = False
.lblAdvanced.Visible = False
.pushID.Visible = False
.ActionInspect.Visible = False
.ActionMove.Visible = False
.ActionNothing.Visible = False
.ActionPath.Visible = False
.Frame1.Visible = False
.chkRuneAlarm.Value = 0
.chkRuneAlarm.enabled = False
.chkRuneAlarm.Visible = False
.txtAlarmUHs.Text = -1
.txtAlarmUHs.enabled = False
.txtAlarmUHs.Visible = False
End With
frmMenu.cmdHardcoreCheats2.enabled = False
#End If
#If KeepRemote Then
With frmHardcoreCheats
.Caption = "Cheats (limited to accept remote orders)"
.lblLeader.Caption = "Only accept order from this leader (locked in this version) :"
.chkAcceptSDorder.Value = save1
.chkAcceptSDorder.enabled = True
.txtRemoteLeader.enabled = False
.chkAcceptSDorder.Visible = True
.txtOrder.Visible = True
.lblOrder2.Visible = True
.lblRead.Visible = True
.cmbOrderType.Visible = True
.lblOn.Visible = True
.lblLeader.Visible = True
.txtRemoteLeader.Visible = True
.chkAcceptSDorder.Top = 100
.txtOrder.Top = 100
.lblOrder2.Top = 100
.lblRead.Top = 340
.cmbOrderType.Top = 340
.lblOn.Top = 340
.lblLeader.Top = 700
.txtRemoteLeader.Top = 680
.txtRemoteLeader.Left = 4500
.txtRemoteLeader.Width = 1000
.Height = 1500
End With
frmMenu.cmdHardcoreCheats2.enabled = True
#End If
End Sub





Private Sub Form_Initialize()
    InitCommonControls
End Sub



Public Function LoadCharSettings(idConnection As Integer, Optional charName As String = "") As String
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    Dim loadCharName As String
    Dim strSettings As String
    Dim pieces() As String
    Dim strLine As String
    Dim ai As Long
    Dim strVarName As String
    Dim strVarValue As String
    Dim posSpliter As Long
    Dim blnTemp As Boolean
    If AutoloadUsable = False Then
        LoadCharSettings = "Autoload is not usable in this environment"
        Exit Function
    End If
    If GameConnected(idConnection) = False Then
        LoadCharSettings = "Character is not connected"
        Exit Function
    End If
    If charName = "" Then
        loadCharName = CharacterName(idConnection)
    Else
        loadCharName = charName
    End If
    strSettings = GetSettingsOfChar(loadCharName)
    If strSettings = "" Then
        LoadCharSettings = "System could not find saved settings found for character " & loadCharName
        Exit Function
    End If
    pieces = Split(strSettings, vbCrLf)
    For ai = 0 To UBound(pieces)
      strLine = Trim$(pieces(ai))
      If strLine <> "" Then
       posSpliter = InStr(1, strLine, "=", vbTextCompare)
       If (posSpliter > 0) Then
        strVarName = Left$(strLine, posSpliter - 1)
        strVarValue = Right$(strLine, Len(strLine) - posSpliter)
        LoadThisCharSetting idConnection, strVarName, strVarValue
       End If
      End If
    Next ai
    LoadCharSettings = ""
    Exit Function
goterr:
    LoadCharSettings = "Unexpected error #" & CStr(Err.Number) & " at LoadCharSettings: " & Err.Description
End Function

Private Sub LoadThisCharSetting(idConnection As Integer, strVar As String, strValue As String)
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    Dim i As Long
    Dim blnTemp As Boolean
    Dim aRes As Long
    Dim tmpStr As String
    Dim tempID As Long
    Dim subValue1 As String
    Dim subValue2 As String
    Dim pieces() As String
    
    'Debug.Print "Loaded:" & strVar & "=" & strValue & "<<<"
    Select Case strVar
    Case "BEGIN_CavebotScript"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmCavebot.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmCavebot.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
    
    
        If blnTemp = True Then
            cavebotIDselected = frmCavebot.cmbCharacter.ListIndex
            cavebotScript(cavebotIDselected).RemoveAll
            cavebotLenght(cavebotIDselected) = 0
            frmCavebot.UpdateValues
        End If
    Case "ADD_CavebotLine"
        AddIDLine cavebotIDselected, cavebotLenght(cavebotIDselected), strValue
        cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) + 1
    Case "END_CavebotScript"
        frmCavebot.UpdateValues
    Case "LastCavebotFile"
        frmCavebot.txtFile.Text = strValue
    Case "CavebotEnabled"
        If strValue = "1" Then
          tmpStr = "exiva openbp"
          tempID = GetTickCount() + 1000
          AddSchedule idConnection, tmpStr, tempID
          frmCavebot.TurnCavebotState idConnection, True
        Else
            frmCavebot.TurnCavebotState idConnection, False
        End If
    Case "BEGIN_Runemaker"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmRunemaker.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmRunemaker.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            runemakerIDselected = frmRunemaker.cmbCharacter.ListIndex
            frmRunemaker.UpdateValues
        End If
    Case "Runemaker_autoEat"
        RuneMakerOptions(runemakerIDselected).autoEat = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutAnyFloor"
        RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutCurrentFloor"
        RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutOutOfRunes"
        RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoWaste"
        RuneMakerOptions(runemakerIDselected).autoWaste = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autossap"
        RuneMakerOptions(runemakerIDselected).autossap = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoerg"
        RuneMakerOptions(runemakerIDselected).autoerg = UnifiedStringToBoolean(strValue)
    Case "Runemaker_firstActionMana"
        RuneMakerOptions(runemakerIDselected).firstActionMana = CLng(strValue)
    Case "Runemaker_beeploot"
        RuneMakerOptions(runemakerIDselected).beeploot = strValue
    Case "Runemaker_text2"
        RuneMakerOptions(runemakerIDselected).Text2 = CLng(strValue)
    Case "Runemaker_text3"
        RuneMakerOptions(runemakerIDselected).Text3 = CLng(strValue)
    Case "Runemaker_firstActionText"
        RuneMakerOptions(runemakerIDselected).firstActionText = strValue
    Case "Runemaker_cmbleaderText"
        RuneMakerOptions(stealthIDselected).cmbleaderText = strValue
    Case "Runemaker_comboText"
        RuneMakerOptions(stealthIDselected).comboText = strValue
    Case "Runemaker_synccomboText"
        RuneMakerOptions(stealthIDselected).synccomboText = strValue
    Case "Runemaker_cmbtypeText"
        RuneMakerOptions(stealthIDselected).cmbtypeText = strValue
    Case "Runemaker_thirdActionText"
        RuneMakerOptions(runemakerIDselected).thirdActionText = CLng(strValue)
    Case "Runemaker_LowMana"
        RuneMakerOptions(runemakerIDselected).LowMana = CLng(strValue)
    Case "Runemaker_ManaFluid"
        RuneMakerOptions(runemakerIDselected).ManaFluid = UnifiedStringToBoolean(strValue)
        If (RuneMakerOptions(runemakerIDselected).ManaFluid = False) Then
            RemoveSpamOrder CInt(runemakerIDselected), 4 'remove auto mana
        End If
    Case "Runemaker_autoUtamo"
        RuneMakerOptions(runemakerIDselected).autoUtamo = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autotar"
        RuneMakerOptions(runemakerIDselected).autotar = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoAp"
        RuneMakerOptions(runemakerIDselected).autoAp = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autossa"
        RuneMakerOptions(runemakerIDselected).autossa = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autopmax"
        RuneMakerOptions(runemakerIDselected).autopmax = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoSdt"
        RuneMakerOptions(runemakerIDselected).autoSdt = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoDan"
        RuneMakerOptions(runemakerIDselected).autoDan = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autodd"
        RuneMakerOptions(runemakerIDselected).autodd = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoee"
        RuneMakerOptions(runemakerIDselected).autoee = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoarme4"
        RuneMakerOptions(runemakerIDselected).autoarme4 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoarme5"
        RuneMakerOptions(runemakerIDselected).autoarme5 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoarme6"
        RuneMakerOptions(runemakerIDselected).autoarme6 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autora"
        RuneMakerOptions(runemakerIDselected).autora = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoda"
        RuneMakerOptions(runemakerIDselected).autoda = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoxray"
        RuneMakerOptions(runemakerIDselected).autoxray = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autodk"
        RuneMakerOptions(runemakerIDselected).autodk = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autogHur"
        RuneMakerOptions(runemakerIDselected).autogHur = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoHur"
        RuneMakerOptions(runemakerIDselected).autoHur = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoPM2"
        RuneMakerOptions(runemakerIDselected).autoPM2 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoaim"
        RuneMakerOptions(runemakerIDselected).autoaim = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoUE"
        RuneMakerOptions(runemakerIDselected).autoUE = UnifiedStringToBoolean(strValue)
    Case "Runemaker_locktrigger"
        RuneMakerOptions(runemakerIDselected).locktrigger = UnifiedStringToBoolean(strValue)
    Case "Runemaker_msgSound"
        RuneMakerOptions(runemakerIDselected).msgSound = UnifiedStringToBoolean(strValue)
    Case "Runemaker_msgSound2"
        RuneMakerOptions(runemakerIDselected).msgSound2 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_secondActionMana"
        RuneMakerOptions(runemakerIDselected).secondActionMana = CLng(strValue)
    Case "Runemaker_secondActionSoulpoints"
        RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = CLng(strValue)
    Case "Runemaker_secondActionText"
        RuneMakerOptions(runemakerIDselected).secondActionText = strValue
    Case "Runemaker_activated"
        RuneMakerOptions(runemakerIDselected).activated = UnifiedStringToBoolean(strValue)
    Case "END_Runemaker"
        frmRunemaker.UpdateValues
        
        'begin hardcore
    Case "BEGIN_HardcoreCheats"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmHardcoreCheats.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmHardcoreCheats.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            HardcoreCheatsIDselected = frmHardcoreCheats.cmbCharacter.ListIndex
            frmHardcoreCheats.UpdateValues
        End If
        
    Case "HardcoreCheats_txtExuraVita3"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita3 = strValue
    
    Case "HardcoreCheats_txtExuraVitaMana2"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana2 = strValue
        
    Case "HardcoreCheats_txtExuraVitaMana"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVitaMana = strValue
    
    Case "HardcoreCheats_txtExuraVita2"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita2 = strValue
        
    Case "HardcoreCheats_Text11"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text11 = strValue
        
    Case "HardcoreCheats_Text10"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text10 = strValue
        
    Case "HardcoreCheats_Text7"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text7 = strValue
        
    Case "HardcoreCheats_Text8"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text8 = strValue
        
    Case "HardcoreCheats_Text2"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text2 = strValue
        
    Case "HardcoreCheats_Text3"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3 = strValue
    
    Case "HardcoreCheats_Text12"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text12 = strValue
        
    Case "HardcoreCheats_Text3"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text3 = strValue
        
    Case "HardcoreCheats_Text6"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text6 = strValue
        
    Case "HardcoreCheats_Text5"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).Text5 = strValue
    
    Case "HardcoreCheats_txtExuraVita4"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita4 = strValue
    
    Case "HardcoreCheats_txtExuraVita"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).txtExuraVita = strValue
        
    Case "HardcoreCheats_arme"
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme = UnifiedStringToBoolean(strValue)
    
    Case "HardcoreCheats_arme2"
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme2 = UnifiedStringToBoolean(strValue)
    
    Case "HardcoreCheats_arme3"
    HardcoreCheatsOptions(HardcoreCheatsIDselected).arme3 = UnifiedStringToBoolean(strValue)
    
    Case "HardcoreCheats_sphi"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).sphi = UnifiedStringToBoolean(strValue)

    Case "HardcoreCheats_splo"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).splo = UnifiedStringToBoolean(strValue)

    Case "HardcoreCheats_pmh"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).pmh = UnifiedStringToBoolean(strValue)

    Case "HardcoreCheats_pth"
        HardcoreCheatsOptions(HardcoreCheatsIDselected).pth = UnifiedStringToBoolean(strValue)

    Case "HardcoreCheats_StopOnGM"
        If strValue = "1" Then
            frmAdvanced.chkStopOnGM.Value = 1
        Else
            frmAdvanced.chkStopOnGM.Value = 0
        End If
    Case "END_HardcoreCheatsr"
        frmHardcoreCheats.UpdateValues
        

    Case "BEGIN_CustomCondEvents"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmCondEvents.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmCondEvents.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            'frmCondEvents.UpdateValues
            condEventsIDselected = frmCondEvents.cmbCharacter.ListIndex
            frmCondEvents.DeleteAllCondEvents CLng(idConnection)
            frmCondEvents.UpdateValues
        End If
    Case "CustomCondEvents_thing1"
        Aux_LastLoadedCond(idConnection).thing1 = strValue
    Case "CustomCondEvents_operator"
        Aux_LastLoadedCond(idConnection).operator = strValue
    Case "CustomCondEvents_thing2"
        Aux_LastLoadedCond(idConnection).thing2 = strValue
    Case "CustomCondEvents_delay"
        Aux_LastLoadedCond(idConnection).delay = strValue
    Case "CustomCondEvents_lock"
        Aux_LastLoadedCond(idConnection).lock = strValue
    Case "CustomCondEvents_keep"
        Aux_LastLoadedCond(idConnection).keep = strValue
    Case "CustomCondEvents_action"
        Aux_LastLoadedCond(idConnection).action = strValue
    Case "CustomCondEvents_ADD"
        aRes = frmCondEvents.AddCondEvent(idConnection, _
         Aux_LastLoadedCond(idConnection).thing1, _
         Aux_LastLoadedCond(idConnection).operator, _
         Aux_LastLoadedCond(idConnection).thing2, _
         Aux_LastLoadedCond(idConnection).delay, _
         Aux_LastLoadedCond(idConnection).lock, _
         Aux_LastLoadedCond(idConnection).keep, _
         Aux_LastLoadedCond(idConnection).action)
    Case "END_CustomCondEvents"
         frmCondEvents.UpdateValues
    Case "BEGIN_Trainer"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmTrainer.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmTrainer.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            'frmCondEvents.UpdateValues
            trainerIDselected = frmTrainer.cmbCharacter.ListIndex
            frmTrainer.UpdateValues
        End If
    Case "Trainer_AllowedSides"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).AllowedSides(CLng(subValue1)) = UnifiedStringToBoolean(subValue2)
    Case "Trainer_idToAvoid"
        TrainerOptions(idConnection).idToAvoid = CLng(strValue)
    Case "Trainer_maxitems"
        TrainerOptions(idConnection).maxitems = CLng(strValue)
    Case "Trainer_misc_avoidID"
        TrainerOptions(idConnection).misc_avoidID = CLng(strValue)
    Case "Trainer_misc_stoplowhp"
        TrainerOptions(idConnection).misc_stoplowhp = CLng(strValue)
    Case "Trainer_spearDest"
        TrainerOptions(idConnection).spearDest = CLng(strValue)
    Case "Trainer_spearID_b1"
        TrainerOptions(idConnection).spearID_b1 = CByte("&H" & strValue)
    Case "Trainer_spearID_b2"
        TrainerOptions(idConnection).spearID_b2 = CByte("&H" & strValue)
    Case "Trainer_stoplowhpHP"
        TrainerOptions(idConnection).stoplowhpHP = CLng(strValue)
    Case "Trainer_PlayerSlots_cheked"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).cheked = CLng(subValue2)
    Case "Trainer_PlayerSlots_itemID_b1"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).itemID_b1 = CByte("&H" & subValue2)
     Case "Trainer_PlayerSlots_itemID_b2"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).itemID_b2 = CByte("&H" & subValue2)
     Case "Trainer_PlayerSlots_xvalue"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).xvalue = CLng(subValue2)
    Case "Trainer_enabled"
        TrainerOptions(idConnection).enabled = CLng(strValue)
    Case "END_Trainer"
      trainerIDselected = idConnection
      frmTrainer.UpdateValues
    End Select
    Exit Sub
goterr:
    Exit Sub
End Sub

