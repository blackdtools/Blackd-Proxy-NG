VERSION 5.00
Begin VB.Form frmStealth 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aimbot"
   ClientHeight    =   3555
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   3270
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmStealth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chklocktrigger 
      Caption         =   "Lock on leader's target"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CheckBox chkautoUE 
      Caption         =   "Active UE combo"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   2820
      Width           =   1575
   End
   Begin VB.TextBox Text_sync 
      Height          =   315
      Left            =   1440
      TabIndex        =   17
      Text            =   "good bye"
      Top             =   2040
      Width           =   1515
   End
   Begin VB.TextBox Text_combo 
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Text            =   "exevo gran mas flam"
      Top             =   1620
      Width           =   1515
   End
   Begin VB.Timer Timeraimbot 
      Interval        =   100
      Left            =   2700
      Top             =   2400
   End
   Begin VB.TextBox Text_cmbleader 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Top             =   780
      Width           =   1515
   End
   Begin VB.TextBox Text_aimtype 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Text            =   "sd"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.CheckBox chkautoaim 
      Caption         =   "Execute Aim type"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CheckBox chkAvoidChat 
      BackColor       =   &H00000000&
      Caption         =   "Avoid chat here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkStealthExp 
      BackColor       =   &H00000000&
      Caption         =   "Exp here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkStealthMessages 
      BackColor       =   &H00000000&
      Caption         =   "Bot messages here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkStealthCommands 
      BackColor       =   &H00000000&
      Caption         =   "Commands here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   7500
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Text            =   "-"
      Top             =   180
      Width           =   1815
   End
   Begin VB.TextBox txtBoard 
      BackColor       =   &H00404040&
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
      Height          =   1575
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Width           =   8415
   End
   Begin VB.Label Label_synccombo 
      Caption         =   "Sync combo :"
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label_spellcombo 
      Caption         =   "Spell Combo :"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000002&
      Index           =   3
      X1              =   60
      X2              =   3180
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000002&
      Index           =   2
      X1              =   60
      X2              =   3180
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000002&
      Index           =   1
      X1              =   3180
      X2              =   3180
      Y1              =   660
      Y2              =   3420
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   60
      X2              =   60
      Y1              =   660
      Y2              =   3420
   End
   Begin VB.Label Label_combo 
      Caption         =   "Aim type :"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label_leader 
      Caption         =   "Leader :"
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblCommand 
      BackColor       =   &H00000000&
      Caption         =   "Command >"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label lblChar 
      BackColor       =   &H80000014&
      Caption         =   "Trigger Aimbot :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmStealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Const WMVSCROLL As Long = &H115
Private Const SBBOTTOM As Long = 7
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const TheLastCommand As Long = 5
Private LastCommand(0 To TheLastCommand) As String
Private CurrCommand As Long
Private previewCommand As Long

'Private Sub Doresize()
'  If frmStealth.WindowState <> vbMinimized Then
'    If frmStealth.ScaleHeight < 2000 Then
'      frmStealth.Height = 2000
'    End If
'    If frmStealth.ScaleWidth < 8100 Then
'      frmStealth.Width = 8100
'    End If
'    txtBoard.Height = frmStealth.ScaleHeight - 800
'    txtBoard.Width = frmStealth.ScaleWidth
'    txtCommand.Top = txtBoard.Height + 465
'    txtCommand.Width = txtBoard.Width - 1250
'    Me.lblCommand.Top = txtCommand.Top
'  End If
'End Sub

Private Sub chkautoaim_Click()
Dim idConnection As Integer

If lock_chkautoaim = False Then
If runemakerIDselected > 0 Then
  If chkautoaim.Value = 1 Then
    RuneMakerOptions(stealthIDselected).autoaim = True
  Else
    RuneMakerOptions(stealthIDselected).autoaim = False
  End If
End If
End If
End Sub

Private Sub chkautoUE_Click()
Dim idConnection As Integer

If lock_chkautoUE = False Then
If runemakerIDselected > 0 Then
  If chkautoUE.Value = 1 Then
    RuneMakerOptions(stealthIDselected).autoUE = True
  Else
    RuneMakerOptions(stealthIDselected).autoUE = False
  End If
End If
End If
End Sub

Private Sub chklocktrigger_Click()
Dim idConnection As Integer

If lock_chklocktrigger = False Then
If runemakerIDselected > 0 Then
  If chklocktrigger.Value = 1 Then
    RuneMakerOptions(stealthIDselected).locktrigger = True
  Else
    RuneMakerOptions(stealthIDselected).locktrigger = False
  End If
End If
End If
End Sub

Private Sub cmbCharacter_Click()
 stealthIDselected = cmbCharacter.ListIndex
  If stealthIDselected = 0 Then
     stealthIDselected = 1
   '   Exit Sub
  End If
 UpdateValues
End Sub

Private Sub Form_Load()
    Dim i As Long
    'Doresize
    LoadStealthChars
    For i = 0 To TheLastCommand
        LastCommand(i) = ""
    Next i
    CurrCommand = TheLastCommand
    previewCommand = TheLastCommand
End Sub

Private Sub Form_Resize()
'Doresize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub UpdateValues()
    '...
    Dim theindex As Integer
    theindex = CInt(cmbCharacter.ListIndex)
    If theindex = 0 Then
        Me.txtBoard.Text = "Select a character so you can read all their bot messages." & vbCrLf & _
        "All commands typed here will be also executed for that character." & vbCrLf & _
        "Commands casted while no character is selected will be ignored."
        Me.Caption = "Aimbot"
    Else
        If Len(stealthLog(theindex)) > 5000 Then
            stealthLog(theindex) = "--Log cleared in order to save memory--"
        End If
        Me.txtBoard.Text = stealthLog(theindex)
        
        ScrollToBottom
    End If
    
    If stealthIDselected = 0 Then
        If chkautoaim = True Then
        chkautoaim.Value = 1
        Else
        chkautoaim.Value = 0
        End If
        If chkautoUE = True Then
        chkautoUE.Value = 1
        Else
        chkautoUE.Value = 0
        End If
        If chklocktrigger = True Then
        chklocktrigger.Value = 1
        Else
        chklocktrigger.Value = 0
        End If
        Text_cmbleader.Text = RuneMakerOptions_cmbleaderText_default
        Text_aimtype.Text = RuneMakerOptions_cmbtypeText_default
        Text_combo.Text = RuneMakerOptions_comboText_default
        Text_sync.Text = RuneMakerOptions_synccomboText_default
    Else
        If chkautoaim = True Then
        chkautoaim.Value = 1
        Else
        chkautoaim.Value = 0
        End If
        If chkautoUE = True Then
        chkautoUE.Value = 1
        Else
        chkautoUE.Value = 0
        End If
        If chklocktrigger = True Then
        chklocktrigger.Value = 1
        Else
        chklocktrigger.Value = 0
        End If
        Text_cmbleader.Text = RuneMakerOptions(stealthIDselected).cmbleaderText
        Text_aimtype.Text = RuneMakerOptions(stealthIDselected).cmbtypeText
        Text_combo.Text = RuneMakerOptions(stealthIDselected).comboText
        Text_sync.Text = RuneMakerOptions(stealthIDselected).synccomboText
    End If
    
    stealthIDselected = theindex
End Sub
Public Sub LoadStealthChars()
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
  stealthIDselected = firstC
  UpdateValues
End Sub

Public Sub ScrollToBottom()
   SendMessage txtBoard.hwnd, WMVSCROLL, SBBOTTOM, 0
End Sub



Private Sub Text_aimtype_Change()
If stealthIDselected > 0 Then
  RuneMakerOptions(stealthIDselected).cmbtypeText = Text_aimtype.Text
End If
End Sub

Private Sub Text_cmbleader_Change()
If stealthIDselected > 0 Then
  RuneMakerOptions(stealthIDselected).cmbleaderText = Text_cmbleader.Text
End If
End Sub

Private Sub Text_combo_Change()
If stealthIDselected > 0 Then
  RuneMakerOptions(stealthIDselected).comboText = Text_combo.Text
End If
End Sub

Private Sub Text_sync_Change()
If stealthIDselected > 0 Then
  RuneMakerOptions(stealthIDselected).synccomboText = Text_sync.Text
End If
End Sub

Private Sub Timeraimbot_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer
Dim aimtype As String
Dim lastmsn As String
Dim UEspell As String
Dim rightpart As String

For idConnection = 1 To MAXCLIENTS
If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then

If RuneMakerOptions(idConnection).cmbtypeText = "sd" Then
aimtype = "exiva 5"
ElseIf RuneMakerOptions(idConnection).cmbtypeText = "hmm" Then
aimtype = "exiva 6"
ElseIf RuneMakerOptions(idConnection).cmbtypeText = "explo" Then
aimtype = "exiva 7"
ElseIf RuneMakerOptions(idConnection).cmbtypeText = "icicle" Then
aimtype = "exiva D:"
Else
aimtype = "exiva 0"
End If

End If
Next idConnection

'lastmsn = LCase(var_lastmsg(idConnection))

    For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then
       
      If (RuneMakerOptions(idConnection).autoaim = True) Then
      lastmsn = LCase(var_lastmsg(idConnection))
        If (var_lastsender(idConnection) = RuneMakerOptions(idConnection).cmbleaderText) Then
                aRes = ExecuteInTibia(aimtype & lastmsn, idConnection, True)
                'var_lastmsg(idConnection) = ""
                If (RuneMakerOptions(idConnection).locktrigger = True) Then
                    currTargetName(idConnection) = lastmsn
                    aRes = ProcessKillOrder2(idConnection, rightpart)
                End If
        End If
      End If
      
      If (RuneMakerOptions(idConnection).autoUE = True) Then
      UEspell = RuneMakerOptions(idConnection).comboText
        If (var_lastsender(idConnection) = RuneMakerOptions(idConnection).cmbleaderText) And (var_lastmsg(idConnection) = RuneMakerOptions(idConnection).synccomboText) Then
                aRes = ExecuteInTibia(UEspell, idConnection, True)
                'var_lastmsg(idConnection) = ""
        End If
      End If
      
      If (RuneMakerOptions(idConnection).autoUE = True) Or (RuneMakerOptions(idConnection).autoaim = True) Then
      var_lastmsg(idConnection) = ""
      End If
      
    End If
    Next idConnection
    
    
End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 38) And (Shift = 1)) Then ' shift + up
        txtCommand.Text = LastCommand(previewCommand)
        previewCommand = previewCommand - 1
        If previewCommand < 0 Then
             previewCommand = TheLastCommand
        End If
  
'        If Len(txtCommand.Text) > 0 Then
'        txtCommand.SelStart = Len(txtCommand.Text) - 1
'        txtCommand.SelLength = 0
'        End If
    ElseIf ((KeyCode = 40) And (Shift = 1)) Then ' shift + down
        txtCommand.Text = LastCommand(previewCommand)
        previewCommand = previewCommand + 1
        If previewCommand > TheLastCommand Then
             previewCommand = 0
        End If
   
'        If Len(txtCommand.Text) > 0 Then
'        txtCommand.SelStart = Len(txtCommand.Text) - 1
'        txtCommand.SelLength = 0
'        End If
    End If
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
    Dim strCommand As String
    Dim iRes As Integer
    If KeyAscii = 13 Then
        strCommand = Trim$(txtCommand.Text)
        If ((txtCommand.Text <> "")) Then
            LastCommand(CurrCommand) = strCommand
            CurrCommand = CurrCommand + 1
            If CurrCommand > TheLastCommand Then
                CurrCommand = 0
            End If
            If stealthIDselected > 0 Then
                If chkAvoidChat.Value = 1 Then
                    iRes = ExecuteInTibia(strCommand, stealthIDselected, True, True)
                Else
                    iRes = ExecuteInTibia(strCommand, stealthIDselected, True)
                End If
            End If
            txtCommand.Text = ""
        End If
    Else
        previewCommand = CurrCommand
    End If

    
End Sub

