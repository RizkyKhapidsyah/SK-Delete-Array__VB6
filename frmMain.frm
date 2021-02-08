VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrays"
   ClientHeight    =   2040
   ClientLeft      =   3540
   ClientTop       =   780
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5385
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Data"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Data"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Only works on dynamic arrays
Dim FirstArray(), DeleteArray()
Private Sub LoadArray()
    'Load Data in the array
    ReDim FirstArray(10)
    For i = 1 To 10
        FirstArray(i) = "Item " & i
    Next
End Sub
Private Sub PrintData()
    frmMain.Cls
    'Display Array Data on form
    For i = 1 To UBound(FirstArray)
        frmMain.Print FirstArray(i)
    Next
End Sub

Private Sub Command1_Click(index As Integer)
    Select Case index
        Case 0
            'Create a second array
            ReDim DeleteArray(0)
            'Initialize Counter
            Dim elecount As Integer
            elecount = 1
            'Cycle through the array
            For i = 1 To UBound(FirstArray)
                'Look for text box data
                'Making both sides of the comparison upper
                'Case removes case sensitivity
                If UCase(FirstArray(i)) <> UCase(Text1) Then
                    'If the text data wasn't found create
                    'a new element in the second array
                    ReDim Preserve DeleteArray(elecount)
                    'Store the data from the first array
                    'In the second array
                    DeleteArray(elecount) = FirstArray(i)
                    'Increment the counter
                    elecount = elecount + 1
                End If
            Next i
            'If the elements in both arrays are the same
            'here the text data wasn't found so we let the
            'user know
            If UBound(FirstArray) = UBound(DeleteArray) Then
                MsgBox Text1 & " Not Found", 48, "Not Found"
                'Highlight the whole string in the text box
                'for ease of use
                Text1.SetFocus
                Text1.SelLength = Len(Text1)
                'Leave here if not found
                Exit Sub
            End If
            'This makes both arrays equal
            FirstArray() = DeleteArray()
            'Print the modified array data
            PrintData
            'Highlight the last digit in the text
            'for ease of use
            Text1.SetFocus
            Text1.SelStart = Len(Text1) - 1
            Text1.SelLength = 1
        Case 1
            'reload and display the array data
            LoadArray
            PrintData
            'Highlight the last digit in the text
            'for ease of use
            Text1.SetFocus
            Text1.SelStart = Len(Text1) - 1
            Text1.SelLength = 1
        Case 2
            End
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'initial load and display of the array data
    LoadArray
    PrintData
    'pick an element at random
    'to load into text box
    
    i = RandomNumber(10, 1)
    Text1.Text = FirstArray(i)
    
End Sub
Private Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  Randomize Timer
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function

