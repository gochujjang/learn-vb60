VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "For Next"
      Height          =   735
      Left            =   8040
      TabIndex        =   7
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do While Loop"
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do Until"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   5640
      Width           =   3495
   End
   Begin VB.ListBox List3 
      Height          =   3570
      Left            =   8040
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Data"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = 1
Do Until x > Val(Text1.Text)
List1.AddItem (x)
x = x + 1
Loop
End Sub

Private Sub Command2_Click()
x = 1
Do While x <= Val(Text1.Text)
List2.AddItem (x)
x = x + 1
Loop
End Sub

Private Sub Command3_Click()
For x = 1 To Val(Text1.Text)
angka = ""
For b = 1 To x
angka = angka & b
Next
List3.AddItem angka
Next x

End Sub
