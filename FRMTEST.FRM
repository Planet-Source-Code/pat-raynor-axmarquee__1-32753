VERSION 5.00
Object = "*\AAXMrquee.vbp"
Begin VB.Form Form1 
   Caption         =   "Marquee Test"
   ClientHeight    =   1776
   ClientLeft      =   2220
   ClientTop       =   6276
   ClientWidth     =   6372
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1776
   ScaleWidth      =   6372
   Begin VB.Frame Frame1 
      Caption         =   "Scroll Mode"
      Height          =   972
      Left            =   30
      TabIndex        =   2
      Top             =   735
      Width           =   1608
      Begin VB.OptionButton Option1 
         Caption         =   "Left to Right"
         Height          =   204
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   552
         Width           =   1392
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Right to Left"
         Height          =   204
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   252
         Value           =   -1  'True
         Width           =   1392
      End
   End
   Begin VB.TextBox txtMsg 
      Height          =   615
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmTest.frx":0000
      Top             =   1080
      Width           =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   585
      Left            =   4980
      TabIndex        =   1
      Top             =   1080
      Width           =   1368
   End
   Begin AXMarqueeCtl.AXMarquee AXMarquee1 
      Height          =   690
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "ActiveX Marquee"
      Top             =   0
      Width           =   6375
      _ExtentX        =   11240
      _ExtentY        =   1207
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message Text:"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  
  With Command1
    If .Caption = "&Start" Then
      .Caption = "&Stop"
      AXMarquee1.Text = txtMsg.Text
      AXMarquee1.Scrolling = True
    Else
      .Caption = "&Start"
      AXMarquee1.Scrolling = False
    End If
  End With

End Sub

Private Sub Form_Load()
  If AXMarquee1.Scrolling Then
    Command1.Caption = "&Stop"
  Else
    Command1.Caption = "&Start"
  End If
  
  Option1(AXMarquee1.ScrollMode).Value = True
End Sub

Private Sub Form_Resize()
  AXMarquee1.Width = Me.ScaleWidth
End Sub

Private Sub Option1_Click(Index As Integer)
  AXMarquee1.ScrollMode = Index
End Sub

