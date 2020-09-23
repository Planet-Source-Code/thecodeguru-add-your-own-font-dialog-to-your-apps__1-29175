VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmsetFont 
   Caption         =   "Font"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Appearance"
      Height          =   1695
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox Check4 
         Caption         =   "Strikethru"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font"
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Font Size"
      Height          =   2535
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample Text"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3135
   End
End
Attribute VB_Name = "frmsetFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Please visit www.computing.iscute.com for more source, tutorials
' e-books and lots of other computing stuffs.
' Please vote for me if you like it. Thanks mate!

Dim ret As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
  Label4.FontBold = True
Else
  Label4.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Label4.FontItalic = True
Else
  Label4.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Label4.FontUnderline = True
Else
  Label4.FontUnderline = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
  Label4.FontStrikethru = True
Else
  Label4.FontStrikethru = False
End If
End Sub

Private Sub Command1_Click()
Form1.Text1.FontName = Label4.FontName
Form1.Text1.FontSize = Label4.FontSize
Form1.Text1.FontBold = Label4.FontBold
Form1.Text1.FontItalic = Label4.FontItalic
Form1.Text1.FontUnderline = Label4.FontUnderline
Form1.Text1.FontStrikethru = Label4.FontStrikethru
Form1.Text1.ForeColor = Label4.ForeColor
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

ret = frm
Debug.Print ret
For x = 1 To Screen.FontCount
List1.AddItem Screen.Fonts(x)
Next
For x = 5 To 72: List2.AddItem Str$(x): Next

For x = 0 To List1.ListCount - 1
 If Form1.Text1.FontName = List1.List(x) Then
  List1.ListIndex = x
  Label4.FontName = List1.List(x)
  Exit For
 End If
Next

For x = 0 To List2.ListCount - 1
 If Int(Val(Form1.Text1.FontSize)) = Val(List2.List(x)) Then
  List2.ListIndex = x
  Label4.FontSize = Val(List2.List(x))
  Text1.Text = List2.List(x)
  Exit For
 End If
Next

If Form1.Text1.FontBold = True Then
 Label4.FontBold = True
 Check1.Value = 1
End If
If Form1.Text1.FontItalic = True Then
 Label4.FontItalic = True
 Check2.Value = 1
End If
If Form1.Text1.FontUnderline = True Then
 Label4.FontUnderline = True
 Check3.Value = 1
End If
If Form1.Text1.FontStrikethru = True Then
 Label4.FontStrikethru = True
 Check4.Value = 1
End If

Label3.BackColor = Form1.Text1.ForeColor
Label4.ForeColor = Form1.Text1.ForeColor

End Sub

Private Sub Label3_Click()
CommonDialog1.ShowColor
Label3.BackColor = CommonDialog1.Color
Label4.ForeColor = CommonDialog1.Color
End Sub

Private Sub List1_Click()
Label4.FontName = List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
Text1.Text = List2.List(List2.ListIndex)
Label4.FontSize = Val(Text1.Text)
End Sub

Private Sub Text1_Change()

For x = 0 To List2.ListCount - 1
If Val(Text1.Text) = Val(List2.List(x)) Then
  List2.ListIndex = x
  Label4.FontSize = Val(Text1.Text)
 Exit For
End If
Next
End Sub
