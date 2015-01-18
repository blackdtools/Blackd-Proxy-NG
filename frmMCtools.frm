VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmMCtools 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MC tools"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   26
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   25
      Top             =   1020
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   24
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   420
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer TimerMCtools 
      Interval        =   100
      Left            =   2940
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manager"
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   3255
      Begin JwldButn2b.JeweledButton cmdnorth 
         Height          =   255
         Left            =   2460
         TabIndex        =   33
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "/\"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin JwldButn2b.JeweledButton cmdright 
         Height          =   255
         Left            =   2820
         TabIndex        =   32
         Top             =   1500
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   ">"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin JwldButn2b.JeweledButton cmdsouth 
         Height          =   255
         Left            =   2460
         TabIndex        =   31
         Top             =   1500
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "\/"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin JwldButn2b.JeweledButton cmdleft 
         Height          =   255
         Left            =   2100
         TabIndex        =   30
         Top             =   1500
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "<"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "turn"
         Height          =   255
         Left            =   1380
         TabIndex        =   29
         Top             =   1500
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "walk"
         Height          =   195
         Left            =   1380
         TabIndex        =   28
         Top             =   1200
         Width           =   795
      End
      Begin JwldButn2b.JeweledButton cmdatk 
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   360
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
         Caption         =   "Atk"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin VB.TextBox Textuecombo 
         Height          =   315
         Left            =   780
         TabIndex        =   21
         Text            =   "exevo gran mas flam"
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox Text_sdcombo 
         Height          =   285
         Left            =   780
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin JwldButn2b.JeweledButton cmdUE 
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Caption         =   "Cast msg"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin JwldButn2b.JeweledButton cmdSDcombo 
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Caption         =   "SD"
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
         BorderColor_Hover=   16761024
         BorderColor_Inner=   16777215
      End
      Begin VB.Label Labelspell 
         Caption         =   "Spell :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Labeltarget 
         Caption         =   "Target :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label LabelMP 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   15
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label LabelMP 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   14
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label LabelMP 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   13
      Top             =   720
      Width           =   675
   End
   Begin VB.Label LabelMP 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   420
      Width           =   675
   End
   Begin VB.Label LabelMP 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   675
   End
   Begin VB.Label LabelHP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1860
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LabelHP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1860
      TabIndex        =   9
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label LabelHP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1860
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label LabelHP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1860
      TabIndex        =   7
      Top             =   420
      Width           =   615
   End
   Begin VB.Label LabelHP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1860
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H80000003&
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   540
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label 
      BackColor       =   &H80000003&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   540
      TabIndex        =   4
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label Label 
      BackColor       =   &H80000003&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   3
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label 
      BackColor       =   &H80000003&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   2
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label 
      BackColor       =   &H80000003&
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmMCtools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdatk_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long
Dim name As String

name = Text_sdcombo.Text

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
        currTargetName(idConnection) = name
        aRes = ProcessKillOrder2(idConnection, currTargetName(idConnection))
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdleft_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
        If Option1.Value = True Then
        aRes = ExecuteInTibia("exiva > 68", idConnection, True)
        Else
        aRes = ExecuteInTibia("exiva turn3", idConnection, True)
        End If
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdnorth_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
        If Option1.Value = True Then
        aRes = ExecuteInTibia("exiva > 65", idConnection, True)
        Else
        aRes = ExecuteInTibia("exiva turn0", idConnection, True)
        End If
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdright_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
        If Option1.Value = True Then
        aRes = ExecuteInTibia("exiva > 66", idConnection, True)
        Else
        aRes = ExecuteInTibia("exiva turn1", idConnection, True)
        End If
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdSDcombo_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long
Dim name As String

name = Text_sdcombo.Text

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
    aRes = ExecuteInTibia("exiva 5" & name, idConnection, True)
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdsouth_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
        If Option1.Value = True Then
        aRes = ExecuteInTibia("exiva > 67", idConnection, True)
        Else
        aRes = ExecuteInTibia("exiva turn2", idConnection, True)
        End If
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub

Private Sub cmdUE_Click()
Dim idConnection As Integer
Dim index As Integer
Dim aRes As Long
Dim name As String

name = Textuecombo.Text

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
    
    If Check1(index).Value = 1 Then
    aRes = ExecuteInTibia(name, idConnection, True)
    End If
    
    index = index + 1

      End If
    Next idConnection
End Sub



Private Sub Option1_Click()

If Option1.Value = True Then
Option2.Value = False
End If

End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
Option1.Value = False
End If

End Sub

Private Sub TimerMCtools_Timer()
Dim idConnection As Integer
Dim index As Integer
'Dim name As String

'name = Text_sdcombo.Text

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       
       Label(index) = CharacterName(idConnection)
       LabelHP(index) = myHP(idConnection)
       LabelMP(index) = myMana(idConnection)
       
        'If Check1(index).Value = 1 And chkholdSD.Value = 1 Then
        'aRes = ExecuteInTibia("exiva 5" & name, idConnection, True)
        'End If
       
       index = index + 1

      End If
    Next idConnection
    
End Sub
