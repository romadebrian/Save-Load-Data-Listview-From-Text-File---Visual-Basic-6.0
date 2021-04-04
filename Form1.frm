VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer

'List1.AddItem Text1.Text
List1.AddItem Date + Time

'pembuatan file Notepad
Open App.Path & "\test.txt" For Append As #1 'Drive penyimpanan

For i = 0 To List1.ListCount - 1
    Print #1, List1.List(i)
Next
Close #1
'MsgBox "Data telah di simpan ke Notepad", 32, "Informasi"
End Sub

Private Sub Load_Click()
Dim ff As Long
Dim line As String

List1.Clear

ff = FreeFile
Open App.Path & "\test.txt" For Input As #ff
Do While Not EOF(ff)
       Line Input #ff, line
       'make sure we're not adding a blank line
       If Len(line) Then List1.AddItem line
Loop
Close #ff
End Sub
