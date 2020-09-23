VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Exploding Form"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "By : Rudy Alex Kohn"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call ExplodeForm(Me, False)
    End
End Sub

Private Sub Form_Load()
    Me.Width = 5000
    Me.Height = 5000
End Sub

Private Sub Form_Resize()
    Label.Left = 0
    Label.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = -1
End Sub

