VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test MS Agent"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpeek 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1395
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   4575
   End
   Begin VB.CommandButton Button 
      Caption         =   "Say it !"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      ToolTipText     =   "Type some text into the text box, then press this button to hear Genie say it"
      Top             =   1860
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_Click()
Agent.Speak txtSpeek.Text
End Sub

Private Sub Form_Load()
txtSpeek.Text = "you are under attack. assess your tactical options."
End Sub
