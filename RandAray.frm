VERSION 5.00
Begin VB.Form frmRndArray 
   AutoRedraw      =   -1  'True
   Caption         =   "Random Object/Data Generator"
   ClientHeight    =   3870
   ClientLeft      =   870
   ClientTop       =   3600
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6210
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      Left            =   3000
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   2835
      TabIndex        =   10
      Top             =   240
      Width           =   2895
      Begin VB.PictureBox Document 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3195
         ScaleWidth      =   2835
         TabIndex        =   11
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox TxtNPLine 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "10"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton CmdRand 
         Caption         =   "Get Random Numbers"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Text            =   "100"
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "Clear Previous Random Numbers"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Numbers Per Line"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "To:"
         Height          =   255
         Index           =   2
         Left            =   1515
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Range of Numbers"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Print Objects"
         Index           =   0
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmRndArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Though technique used in this example as a random number generator the code with little
'modification can be used for a variety of purposes:
'       Random object generator (pictures for games puzzles, cards etc)
'       Random name generator (Picking names for exchanging gifts)
'       Random file selector (graphic files in a directory for screensaver slide show)
'       Random number generator (Lotto numbers)
'Basically it randomizes anything that can be stored in an array! The objects/data within
'the array is randomized by randomally selecting elements of the array
'
'There is a lot of cool functions this code does:
'       It Shows how to view and print the data generated (placing picture boxes within
'       Picture boxes and munipulating with scroll bars
'       It randomizes a timer for true random seed
'       It uses a second array to insure one time selection
'       It dynamically adjusts the documents height and width to accommodate the data
'       It randomizes a range(user input) of objects
'       It dynamically formats the data columns to user input
Option Explicit
Dim NumberArray(), DeleteArray(), I As Integer
Dim RN As Single, ANum As Single

'Procedure for setting the values of a horizontal and vertical Scroll bar to the
'size of two picture boxes one within another. It should be called prior to first
'display and then when resizing the document picture box
'Syntax SizeScrolls(Horizontal Scroll Bar, Vertical Scroll Bar, Print Picture box, Container Picture box)
'This is a sample of an independent procedure, it is passed the four objects
'horizontal and vertical scroll bars and the two picture boxes. This independence
'allows the procedure to be copied into any project having those four controls
'and used without modifications!
Public Sub SizeScrolls(H As Object, V As Object, Doc As Object, DocContainer As Object)
    With V
        .Left = DocContainer.Left + DocContainer.Width
        .Top = DocContainer.Top
        .Max = Doc.Height - DocContainer.ScaleHeight '32,767
        .Min = 0
        .Value = .Min
        .Height = DocContainer.Height
        .SmallChange = DocContainer.Height / 10    '1/10 of the container height
        .LargeChange = DocContainer.Height
    End With
    If Doc.ScaleHeight > DocContainer.ScaleHeight Then
        V.Visible = True
    Else
        V.Visible = False
    End If
    With H
        .Left = DocContainer.Left
        .Top = DocContainer.Top + DocContainer.Height
        .Min = 0
        .Width = DocContainer.Width
        .Value = .Min
        .Max = Doc.Width - DocContainer.ScaleWidth
        .SmallChange = DocContainer.ScaleWidth / 10
        .LargeChange = DocContainer.Width
    End With
    If Doc.ScaleWidth > DocContainer.ScaleWidth Then
        H.Visible = True
    Else
        H.Visible = False
    End If
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Select Case Index
        Case 0
            Printer.PaintPicture Document.Image, 0, 0
            Printer.EndDoc
        Case 9
            End
    End Select
End Sub

Private Sub VScroll1_Change()
    Document.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub HScroll1_Change()
    Document.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

'Add the objects/Data to the array in this Procedure
Private Sub ReSetObjects(Max As Integer)
    On Error GoTo ResetErrors
    'ReDimension the array
    ReDim NumberArray(1 To Max)
    'Store the first object (in this case the low number)
    ANum = txtRange(0)
    For I = 1 To UBound(NumberArray)
        'Add the object
        NumberArray(I) = ANum
        'Incremented just for number
        ANum = ANum + 1
    Next
Exit Sub

ResetErrors:
    Select Case Err
        Case 9
            Exit Sub
    End Select
End Sub
Private Function RandomNumber() As Integer
    'Seed the randomizer from the timer
    Randomize Timer
    'Pick a random element number within the array
    RandomNumber = Int(Rnd * UBound(NumberArray) + 1)
End Function

Private Sub CmdRand_Click()
    Dim R As Single
    On Error GoTo CmdRandErrors
    ReSetObjects (txtRange(1) - txtRange(0)) + 1
    If ChkClear Then
        Document.Width = Container.Width
        Document.Height = Container.Height
        Document.Cls
    Else
        If Document.CurrentY > 0 Then Document.Print
    End If
    For R = 1 To UBound(NumberArray)
        RN = 0
        'Get a Random number and display it in the label
        Do While RN = 0
            RN = NumberArray(RandomNumber)
            If RN = 0 Then Exit Do
        Loop
        'Create a second array
        ReDim DeleteArray(0)
        'Initialize Counter
        Dim elecount As Integer
        elecount = 1
        'Cycle through the array
        For I = 1 To UBound(NumberArray)
            'Look for text box data
            'Making both sides of the comparison upper
            'Case removes case sensitivity
            If NumberArray(I) <> RN Then
                'If the text data wasn't found create
                'a new element in the second array
                ReDim Preserve DeleteArray(elecount)
                'Store the data from the first array
                'In the second array
                DeleteArray(elecount) = NumberArray(I)
                'Increment the counter
                elecount = elecount + 1
            End If
        Next I
        'This makes both arrays equal
        NumberArray() = DeleteArray()
        'Formats and prints data to users choice of numbers on a line
        If R Mod TxtNPLine = 0 Then
            If Document.Height < Document.Height + Document.TextHeight(RN) Then
                Document.Height = Document.CurrentY + (Document.TextHeight(RN) * 3)
                DoEvents
                SizeScrolls HScroll1, VScroll1, Document, Container
            End If
            Document.Print RN
        Else
            If Document.Width < Document.CurrentX + (Document.TextWidth(RN) * 4) Then
                Document.Width = Document.CurrentX + (Document.TextWidth(RN) * 4)
                DoEvents
                SizeScrolls HScroll1, VScroll1, Document, Container
            End If
            Document.Print RN & ", ";
        End If
    Next
    'Display last
    Print NumberArray(UBound(NumberArray))
Exit Sub

CmdRandErrors:
    Select Case Err
        Case 9
            Exit Sub
    End Select

End Sub
