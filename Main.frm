VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3360
      Picture         =   "Main.frx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   4
      Top             =   2100
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FindPath"
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   60
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   60
      Width           =   3030
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1635
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Do
        DoEvents
        For a = 1 To 400000
        Next a
        
        MovePlayer
        PaintBoard
        If P1.Path = "" Then
newCoor:
            tx = Int(Rnd * 20) + 1
            ty = Int(Rnd * 20) + 1
            If Board(tx, ty) Then GoTo newCoor
            Tr.X = tx
            Tr.Y = ty
            P1.Path = FindPath(P1.X, P1.Y, Tr.X, Tr.Y)
        End If
    Loop
End Sub

Private Sub Command2_Click()
    PaintBoard
    P1.Path = FindPath(P1.X, P1.Y, Tr.X, Tr.Y)
End Sub

Private Sub Form_Load()
    Randomize
    LoadBoard
    P1.X = 9
    P1.Y = 2
    Tr.X = 8
    Tr.Y = 4
    PaintBoard
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tx = Int(X / Size) + 1
    ty = Int(Y / Size) + 1
    If Not Button <> 1 Then
        If Board(tx, ty) = False Then
            P1.X = tx
            P1.Y = ty
            FindPath P1.X, P1.Y, Tr.X, Tr.Y
        End If
    Else
        If Board(tx, ty) = False Then
            Tr.X = tx
            Tr.Y = ty
            FindPath P1.X, P1.Y, Tr.X, Tr.Y
        End If
    End If
    PaintBoard
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tx = Int(X / Size) + 1
    ty = Int(Y / Size) + 1
    Label1.Caption = tx & " - " & ty
End Sub
