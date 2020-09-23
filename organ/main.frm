VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Org"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   2730
   ClientWidth     =   11340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   11340
   Begin VB.Label Label13 
      Caption         =   "to see what key is  belong to what note please see the left table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6840
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      Caption         =   "A2b"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "F7"
      Height          =   255
      Left            =   10920
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   30
      Left            =   10920
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   29
      Left            =   10560
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   28
      Left            =   10200
      Top             =   2640
      Width           =   270
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "B6"
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   13
      Left            =   9720
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   12
      Left            =   9360
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   27
      Left            =   9840
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   26
      Left            =   9480
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   11
      Left            =   8640
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   10
      Left            =   8280
      Top             =   2640
      Width           =   165
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "F6"
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "C6"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   9
      Left            =   7920
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   8
      Left            =   7200
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   7
      Left            =   6840
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   6
      Left            =   6120
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   5
      Left            =   5760
      Top             =   2640
      Width           =   165
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "G4"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "G3"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   4
      Left            =   5400
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   3
      Left            =   4680
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   2
      Left            =   4320
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   1
      Left            =   3240
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   975
      Index           =   0
      Left            =   2880
      Top             =   2640
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   25
      Left            =   9120
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   24
      Left            =   8760
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   23
      Left            =   8400
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   22
      Left            =   8040
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   21
      Left            =   7680
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   20
      Left            =   7320
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   19
      Left            =   6960
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   18
      Left            =   6600
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   17
      Left            =   6240
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   16
      Left            =   5880
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   15
      Left            =   5520
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   14
      Left            =   5160
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   13
      Left            =   4800
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   12
      Left            =   4440
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   11
      Left            =   4080
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   10
      Left            =   3720
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   9
      Left            =   3360
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   8
      Left            =   3000
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   7
      Left            =   2640
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   6
      Left            =   2280
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   5
      Left            =   1920
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   4
      Left            =   1560
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   3
      Left            =   1200
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   2
      Left            =   840
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   1
      Left            =   480
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Index           =   0
      Left            =   120
      Top             =   2640
      Width           =   270
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF80FF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tstart As Single
Dim tend As Single
Dim up As Boolean
Dim note&
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo er1
Dim k As Integer
Dim kp As Integer
k = 500
kp = 500
If (KeyCode = vbKeyA Or KeyCode = vbKeyS Or KeyCode = vbKeyD Or KeyCode = vbKeyF Or KeyCode = vbKeyG Or KeyCode = vbKeyH Or KeyCode = vbKeyJ Or KeyCode = vbKeyK Or KeyCode = vbKeyL Or KeyCode = 222 Or KeyCode = 190 Or KeyCode = vbKeyM Or KeyCode = vbKeyV Or KeyCode = vbKeyX) Then
Select Case KeyCode
Case vbKeyA
kp = 0
Case vbKeyS
kp = 1
Case vbKeyD
kp = 2
Case vbKeyF
kp = 3
Case vbKeyG
kp = 4
Case vbKeyH
kp = 5
Case vbKeyJ
kp = 6
Case vbKeyK
kp = 7
Case vbKeyL
kp = 8
Case 222
kp = 9
Case 190
kp = 10
Case vbKeyM
kp = 11
Case vbKeyV
kp = 12
Case vbKeyX
kp = 13
End Select
Shape2(kp).BackColor = RGB(150, 150, 150) 'change the color of white key to dark gray
Else
Select Case KeyCode
Case vbKey5
k = 0
Case vbKey6
k = 1
Case vbKey7
k = 2
Case vbKey8
k = 3
Case vbKey9
k = 4
Case vbKey0
k = 5
Case 189
k = 6
Case 187
k = 7
Case vbKeyQ
k = 8
Case vbKeyW
k = 9
Case vbKeyE
k = 10
Case vbKeyR
k = 11
 
Case vbKeyT
k = 12
Case vbKeyY
k = 13
Case vbKeyU
k = 14
Case vbKeyI
k = 15
Case vbKeyO
k = 16
Case vbKeyP
k = 17
Case 219
k = 18
Case 221
k = 19
Case 186
k = 20
Case 191
k = 21
Case 188
k = 22
Case vbKeyN
k = 23
Case vbKeyB
k = 24
Case vbKeyC
k = 25
Case vbKeyZ
k = 26
Case vbKey1
k = 27
Case vbKey2
k = 28
Case vbKey3
k = 29
Case vbKey4
k = 30
 
End Select
Shape1(k).BackColor = RGB(100, 100, 100) 'chane the color of black key to gray
End If
er1:
Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo er
Dim clockticks&
Dim loopcount&
Dim z
Dim i As Single
tstart = Timer
Select Case Chr(KeyAscii) '
 
Case "1"
note& = 1396.9 'F6
Case "2"
note& = 1568#  'G6
Case "3"
note& = 1760#  'A6
Case "4"
note& = 1975.5 'B6
Case "5"
note& = 98    'G3
Case "6"
note& = 110   'A3
Case "7"
note& = 123.47 'B3
Case "8"
note& = 130.81 'C4
Case "9"
note& = 146.83 'D4
Case "0"
note& = 164.81 'E4
Case "-"
note& = 174.61 'F4
Case "="
note& = 196    'G4

Case "q"
note& = 220   'A4
Case "w"
note& = 246.94 'B4
Case "e"
note& = 261.63 'C5
Case "r"
note& = 293.66 'D5
Case "t"
note& = 329.63 'E5
Case "y"
note& = 349.23 'F5
Case "u"
note& = 392    'G5
Case "i"
note& = 440    'A5
Case "o"
note& = 493.88 'B5
Case "p"
note& = 523.25 'middle C6
Case "["
note& = 587.33 'D6
Case "]"
note& = 659.26 'E6
Case "a"
note& = 207.65 'A4b
Case "s"
note& = 233.08 'B4b
Case "d"
note& = 277.18 'D5b
Case "f"
note& = 311.13 'E5b
Case "g"
note& = 369.99 'G5b
Case "h"
note& = 415.31 'A5b
Case "j"
note& = 466.16 'B5b
Case "k"
note& = 554.37 'D6b
Case "l"
note& = 622.25 'E6b
Case ";"
note& = 698.46 'F6
Case "'"
note& = 733.99 'G6b


Case "/"
note& = 783.99 'G6
Case "."
note& = 830.61    'A6b
Case ","
note& = 880    'A6
Case "m"
note& = 932.33 'B6b
Case "n"
note& = 987.77 'B6
Case "b"
note& = 1046.5  'C7
Case "v"
note& = 1108.7  'D7b
Case "c"
note& = 1174.7  'D7
Case "x"
note& = 1244.5  'E7b
Case "z"
note& = 1318.5  'E7
End Select
clockticks& = CLng(1193280 \ note&) 'calculate clock ticks
Out 67, 182 'prepare for data
z = Inp(97) 'initial position
Out 66, clockticks& And &HFF 'send data
Out 66, clockticks& \ 256 'send data

Out 97, Inp(97) Or &H3 'turn on speaker
For i = 1 To 50000 '
DoEvents
Next i
Out 97, Inp(97) And &HFC 'turn of speaker (also you can use  Out 97, z)
er:
Exit Sub
End Sub
Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) 'change the color to defult
Dim z As Single
Dim ii As Integer, it As Integer
For ii = 0 To 30
Shape1(ii).BackColor = &H8000000E
Next ii
For it = 0 To 13
Shape2(it).BackColor = &H80000008
Next it
tend = Timer
End Sub
Public Sub Form_Load()
Label1.Caption = " 5=G3   6=A3   7=B3   8=C4   9=D4   0=E4   -=F4     = =G4   "
Label2.Caption = "q=A4  w=B4  e=C5  r=D5  t=E5  y=F5  u=G5  i=A5 o=B5  p=middle C6  [=D6  ]=E6"
Label3.Caption = "a=A4b  s=B4b  d=D5b  f=E5b  g=G5b  h=A5b  j=B5b  k=D6b  l=E6b  ;=F6 '=G6b"
Label4.Caption = " z=E7 x=E7b c=D7  v=D7b b=C7  n=B6  m=B6b ,=A6 .=A6b /=G6 "
Label5.Caption = "1=F7 2=G7 3=A7 4=B7"
End Sub


