VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadListView(FilePath As String, ListView As ListView)

    On Error GoTo Err: 'Our error reporting
    'Our Variables
    Dim Lst As ListItem 'For out ListView
    Dim Fnum As Integer 'For our FreeFile
    Dim tData As String 'Stores Data from a text file
    Dim tAry As Variant 'Stores column items from text file
    Dim F As Integer 'Used With our For Loop
    Const Delim As String = "," 'Our unique delimiter. Feel free To change this If needed. Just watch out For what you use, it can cause problems.
    'Set our ListView to Report Viewing
    ListView.View = lvwReport
    'Let's clear our ListView incase there i
    '     s data already present
    ListView.ListItems.Clear
    'Let's get a free file handle
    Fnum = FreeFile
    'Open our text file for inputing
    Open FilePath For Input As Fnum
    'Do loop while were not at the End Of Fi
    '     le


    Do While Not EOF(Fnum)
        'Input 1 line from our text file into th
        '     e variable tData
        Input #Fnum, tData
        'Split our line of text and store each i
        '     ndvidual value into an array
        tAry = Split(tData, Delim)
        'Loop through each element of the array
        '     and add it to the ListView


        For F = LBound(tAry) To UBound(tAry)
            'If this is the first item


            If F = 0 Then
                'Add our item to the ListView
                Set Lst = ListView.ListItems.Add(, , tAry(F))
            Else
                'Else we need to add the approptiate sub
                '     item
                Lst.SubItems(F) = tAry(F)
            End If

        Next F 'Continue on

    Loop 'Continue on

    Close Fnum 'Close our file
    Exit Sub 'Exit this sub
Err:     ' Our Error reporting. Feel free To fix this up
    MsgBox "Error Locating Data File.", vbCritical
End Sub

Private Sub Command1_Click()
LoadListView "C:\Test.txt", ListView1
End Sub

Private Sub Form_Load()

End Sub
