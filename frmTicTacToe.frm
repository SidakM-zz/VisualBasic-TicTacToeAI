VERSION 5.00
Begin VB.Form frmTicTacToe 
   BackColor       =   &H80000007&
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "I STILL WANT TO TRY :("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   600
      MaskColor       =   &H000080FF&
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00000000&
      Caption         =   "YOU CANNOT WIN THIS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPLay_Click()
frmTicTacHuman.Show 'loads the game form
Unload Me 'unloads this form
End Sub

