VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please visit www.computing.iscute.com"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Set Font"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmsetFont.Show

End Sub
