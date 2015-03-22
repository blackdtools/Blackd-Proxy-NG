VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmBigText 
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Cavebot Script"
   ClientHeight    =   3570
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   7440
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBigText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton cmdCancel 
      Height          =   315
      Left            =   1740
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   "Cancel"
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
   Begin JwldButn2b.JeweledButton cmdClear 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   "Clean"
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
   Begin JwldButn2b.JeweledButton cmdOk 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      Caption         =   "Load Script"
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
   Begin VB.CommandButton cmdClear2 
      BackColor       =   &H80000014&
      Caption         =   "Clean"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel2 
      BackColor       =   &H80000014&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk2 
      BackColor       =   &H80000014&
      Caption         =   "Load Script"
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtBoard 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   7215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "Text board"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2580
      TabIndex        =   4
      Top             =   5760
      Width           =   8055
   End
End
Attribute VB_Name = "frmBigText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Private Sub cmdCancel_Click()
  CanceledBoard = True
  ClosedBoard = True
  frmBigText.Hide
End Sub

Private Sub cmdClear_Click()
  txtBoard.Text = ""
End Sub

Private Sub cmdOk_Click()
  CanceledBoard = False
  ClosedBoard = True
  frmBigText.Hide
End Sub

Private Sub Form_Resize()
  If frmBigText.WindowState <> vbMinimized Then
    If frmBigText.ScaleHeight < 3000 Then
      frmBigText.Height = 3000
    End If
    If frmBigText.ScaleWidth < 5800 Then
      frmBigText.Width = 5800
    End If
    txtBoard.Height = frmBigText.ScaleHeight - 1300
    txtBoard.Width = frmBigText.ScaleWidth - 200
    cmdClear2.Top = frmBigText.ScaleHeight - 480
    cmdOk2.Top = frmBigText.ScaleHeight - 480
    cmdCancel2.Top = frmBigText.ScaleHeight - 480
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CanceledBoard = True
  ClosedBoard = True
  Me.Hide
  Cancel = BlockUnload
End Sub


