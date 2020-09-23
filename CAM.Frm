VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cellular Automata Machine"
   ClientHeight    =   5940
   ClientLeft      =   3015
   ClientTop       =   1635
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   6810
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Updating"
      Height          =   615
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   1080
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "CAM.frx":0000
      Left            =   5040
      List            =   "CAM.frx":0002
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Text            =   "50"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "50"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   1200
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Updating"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Start"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "rules:"
      Height          =   195
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "column"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "row"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim RIGHT As Integer
Dim BOTTOM As Integer
Dim GRID() As Integer
Dim NEWARRAY() As Integer
Dim Rules As String

Function UpdateRule(code As String) As Integer
' Function that takes a window (3x3) as input
' and produce result (dead or alive) on the center
' of the window.
' It is the rules look-up table.
Dim NW, N, NE, W, C, E, SW, S, SE
Dim sum8 As Integer, sum9 As Integer
Dim newCenter As Integer
NW = Val(Mid(code, 1, 1))
N = Val(Mid(code, 2, 1))
NE = Val(Mid(code, 3, 1))
W = Val(Mid(code, 4, 1))
E = Val(Mid(code, 5, 1))
SW = Val(Mid(code, 6, 1))
S = Val(Mid(code, 7, 1))
SE = Val(Mid(code, 8, 1))
C = Val(Mid(code, 9, 1))
sum8 = NW + N + NE + W + E + SW + S + SE
sum9 = NW + N + NE + W + E + SW + S + SE + C
'rules
Select Case Rules
Case "square"
    newCenter = NW Or N Or NE Or W Or E Or SW Or S Or SE + C
Case "diamond"
    newCenter = N Or W Or E Or S
Case "triangle"
    newCenter = N Or W Or E
Case "1 out of 8"
    If sum8 = 1 Then
        newCenter = 1
    Else
        newCenter = C
    End If
Case "Lichen"
    If sum8 = 3 Or sum8 = 7 Or sum8 = 8 Then
        newCenter = 1
    Else
        newCenter = C
    End If
Case "Lichen with Dead"
    If sum8 = 3 Or sum8 = 7 Or sum8 = 8 Then
        newCenter = 1
    ElseIf sum8 = 4 Then
        newCenter = 0
    Else
        newCenter = C
    End If
Case "Anneal"
    If sum9 = 4 Or sum9 >= 6 Then
        newCenter = 1
    Else
        newCenter = 0
    End If
Case "Majority"
    If sum9 >= 5 Then
        newCenter = 1
    Else
        newCenter = 0
    End If
Case "Bank"
    If N = 1 And C = 1 And E = 1 And W = 0 And S = 0 Then
        newCenter = 0
    ElseIf N = 0 And C = 0 And E = 1 And W = 1 And S = 1 Then
        newCenter = 1
    ElseIf N = 1 And C = 0 And E = 1 And W = 1 And S = 1 Then
        newCenter = 1
    Else
        newCenter = C
    End If
Case "Life"
    Select Case C
    Case 1
        If sum8 = 2 Or sum8 = 3 Then
            newCenter = 1
        Else
            newCenter = 0
        End If
    Case 0
        If sum8 = 3 Then
            newCenter = 1
        Else
            newCenter = 0
        End If
    End Select
End Select

UpdateRule = newCenter
End Function

Sub Draw(GRID)
' draw the cell based on current state.
Dim row As Integer, col As Integer
Dim X1, Y1, X2, Y2, Color


For row = 1 To BOTTOM
    For col = 1 To RIGHT
        X1 = (col - 1) * Picture1.Width / RIGHT
        Y1 = (row - 1) * Picture1.Height / BOTTOM
        X2 = col * Picture1.Width / RIGHT
        Y2 = row * Picture1.Height / BOTTOM
        
        If GRID(row, col) = 1 Then
        Color = QBColor(0) '12
        Else
        Color = QBColor(15) '15
        End If
        Picture1.Line (X1, Y1)-(X2, Y2), Color, BF
    Next col
Next row


End Sub
Private Sub Command1_Click()
'put the first points in the PictureBox
Dim row As Integer, col As Integer
Dim X As Integer, Y As Integer
For row = 1 To BOTTOM
    For col = 1 To RIGHT
        GRID(row, col) = Int((1 - 0 + 1) * Rnd + 0) 'random of 0 or 1
    Next col
Next row
Call Draw(GRID)
End Sub


Private Sub Command2_Click()

Timer1.Enabled = True
Command3.Enabled = True
Command2.Enabled = False
End Sub

Function GetWindow(cc() As Integer) As String
' input : center coordinate
' output: code as string in order:
' NW-N-NE-W-E-SW-S-SE-C
Dim X, Y, NW, N, NE, W, C, E, SW, S, SE
Dim code As String

X = cc(1)    'X-th row
Y = cc(2)    'Y-th column

'if corner
Select Case X
Case 1         ' top
    Select Case Y
    Case 1     ' top-left
        NW = GRID(X + 1, Y + 1): N = GRID(X + 1, Y): NE = GRID(X + 1, Y + 1)
        W = GRID(X, Y + 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X + 1, Y + 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y + 1)
    Case RIGHT 'top-right
        NW = GRID(X + 1, Y - 1): N = GRID(X + 1, Y): NE = GRID(X + 1, Y - 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y - 1)
        SW = GRID(X + 1, Y - 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y - 1)
    Case Else  ' top-middle
        NW = GRID(X + 1, Y - 1): N = GRID(X + 1, Y): NE = GRID(X + 1, Y + 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X + 1, Y - 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y + 1)
    End Select
Case BOTTOM
    Select Case Y
    Case 1    'bottom-left
        NW = GRID(X - 1, Y + 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y + 1)
        W = GRID(X, Y + 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X - 1, Y + 1): S = GRID(X - 1, Y): SE = GRID(X - 1, Y + 1)
    Case RIGHT 'bottom-right
        NW = GRID(X - 1, Y - 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y - 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y - 1)
        SW = GRID(X - 1, Y - 1): S = GRID(X - 1, Y): SE = GRID(X - 1, Y - 1)
    Case Else  'bottom-middle
        NW = GRID(X - 1, Y - 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y + 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X - 1, Y - 1): S = GRID(X - 1, Y): SE = GRID(X - 1, Y + 1)
    End Select
Case Else 'middle
    Select Case Y
    Case 1     'middle-left
        NW = GRID(X - 1, Y + 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y + 1)
        W = GRID(X, Y + 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X + 1, Y + 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y + 1)
    Case RIGHT 'middle-right
        NW = GRID(X - 1, Y - 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y - 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y - 1)
        SW = GRID(X + 1, Y - 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y - 1)
    Case Else  'middle-middle (not corner)
        NW = GRID(X - 1, Y - 1): N = GRID(X - 1, Y): NE = GRID(X - 1, Y + 1)
        W = GRID(X, Y - 1): C = GRID(X, Y): E = GRID(X, Y + 1)
        SW = GRID(X + 1, Y - 1): S = GRID(X + 1, Y): SE = GRID(X + 1, Y + 1)
    End Select
End Select

code = Trim(Str(NW)) + Trim(Str(N)) + Trim(Str(NE)) + Trim(Str(W)) + Trim(Str(E)) + Trim(Str(SW)) + Trim(Str(S)) + Trim(Str(SE)) + Trim(Str(C))
GetWindow = code
End Function





Private Sub Command3_Click()
Timer1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Form_Load()
RIGHT = Val(Text1.Text)
BOTTOM = Val(Text2.Text)
ReDim GRID(BOTTOM, RIGHT) As Integer
ReDim NEWARRAY(BOTTOM, RIGHT) As Integer
List1.List(0) = "square"
List1.List(1) = "triangle"
List1.List(2) = "diamond"
List1.List(3) = "1 out of 8"
List1.List(4) = "lichen"
List1.List(5) = "lichen with die"
List1.List(6) = "anneal"
List1.List(7) = "Majority"
List1.List(8) = "Bank"
List1.List(9) = "Life"

List1.ListIndex = 9
Command3.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub List1_Click()
Rules = List1.Text
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim row, col
row = Fix(Y * BOTTOM / Height) + 1
col = Fix(X * RIGHT / Width) + 1
GRID(row, col) = 1
Draw (GRID)
End Sub




Private Sub Text1_Change()
RIGHT = Val(Text1.Text)
If RIGHT < 1 Then RIGHT = 1
ReDim GRID(BOTTOM, RIGHT) As Integer
ReDim NEWARRAY(BOTTOM, RIGHT) As Integer
End Sub


Private Sub Text2_Change()
BOTTOM = Val(Text2.Text)
If BOTTOM < 1 Then BOTTOM = 1
ReDim GRID(BOTTOM, RIGHT) As Integer
ReDim NEWARRAY(BOTTOM, RIGHT) As Integer
End Sub


Private Sub Timer1_Timer()
Dim row As Integer, col As Integer
Dim centerCoord(2) As Integer
Dim code As String
'update state
For row = 1 To BOTTOM
    For col = 1 To RIGHT
        centerCoord(1) = row
        centerCoord(2) = col
        code = GetWindow(centerCoord)
        NEWARRAY(row, col) = UpdateRule(code)
    Next col
Next row

For row = 1 To BOTTOM
    For col = 1 To RIGHT
        GRID(row, col) = NEWARRAY(row, col)
    Next col
Next row

Call Draw(GRID)
End Sub


