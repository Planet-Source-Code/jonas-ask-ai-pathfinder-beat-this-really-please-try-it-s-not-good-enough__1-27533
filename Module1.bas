Attribute VB_Name = "Module1"
Public Type Unit
 X As Integer
 Y As Integer
 Path As String
End Type

Public Board(1 To 20, 1 To 20) As Boolean
Public Const Size As Byte = 10
Public Const HG As Byte = 20
Public Const BR As Byte = 20

Public P1 As Unit
Public Tr As Unit
Dim FinalPath As String

Public AIBoard() As Boolean
Public Sub LoadBoard()
For Y = 1 To HG
        For X = 1 To BR
            If Main.PicMap.Point(X - 1, Y - 1) = 0 Then Board(X, Y) = True
        Next X
    Next Y
End Sub
Public Sub PaintBoard()
    Main.Pic.Cls
    For Y = 1 To HG
        For X = 1 To BR
            If Board(X, Y) Then Main.Pic.Line (X * Size, Y * Size)-Step(-Size, -Size), , BF
        Next X
    Next Y
    Main.Pic.Line (P1.X * Size, P1.Y * Size)-Step(-Size, -Size), vbBlue, BF
    Main.Pic.Line (P1.X * Size, P1.Y * Size)-Step(-Size, -Size), vbBlue, BF
    Main.Pic.Line (Tr.X * Size, Tr.Y * Size)-Step(-Size, -Size), vbGreen, BF
    Main.Pic.Line (Tr.X * Size, Tr.Y * Size)-Step(-Size, -Size), vbGreen, BF
End Sub
Public Function FindPath(sX, sY, eX, eY) As String
Dim LastPathReally As String 'the final path, really, I swear ;)
    For a = 1 To 7 'Find many routs, then get the shortest
        ReDim AIBoard(1 To BR, 1 To HG) 'Clear out the AiBoard
        FinalPath = "" 'empty the path string
        OnRoute sX, sY, eX, eY, "" 'Start genetrating a route
        If LastPathReally = "" Then LastPathReally = FinalPath 'if this is the first path store it
        If Len(FinalPath) > 0 And Len(FinalPath) < Len(LastPathReally) Then LastPathReally = FinalPath 'if the aquierd path is shorter than the current, store it
        FinalPath = "" 'empty once more to not leave anything
    Next a
    FindPath = LastPathReally 'apply the found path
    LastPathReally = ""
End Function
Public Function OnRoute(X, Y, gX, gY, PathSoFar) As String
Dim Checked(1 To 4) As Boolean
NewDire:
    If FinalPath <> "" Then Exit Function
    If Checked(1) And Checked(2) And Checked(3) And Checked(4) Then
        OnRoute = PathSoFar
        Exit Function
    End If
    a = Int(Rnd * 4) + 1
    If Checked(a) Then GoTo NewDire:
    'a = 1
    Select Case a
    Case 1
        Checked(a) = True
        If Movable(X - 1, Y) Then
            PathSoFar = PathSoFar & "l"
            If X - 1 = gX And Y = gY Then GoTo FoundRoute
            AIBoard(X, Y) = True
            OnRoute = OnRoute(X - 1, Y, gX, gY, PathSoFar)
            If OnRoute = PathSoFar Then 'No way found
                OnRoute = Left(OnRoute, Len(OnRoute) - 1)
                PathSoFar = OnRoute
                GoTo NewDire
            End If
        Else: GoTo NewDire
        End If
    Case 2
        Checked(a) = True
        If Movable(X, Y + 1) Then
            PathSoFar = PathSoFar & "d"
            If X = gX And Y + 1 = gY Then GoTo FoundRoute
            AIBoard(X, Y) = True
            OnRoute = OnRoute(X, Y + 1, gX, gY, PathSoFar)
            If OnRoute = PathSoFar Then 'No way found
                OnRoute = Left(OnRoute, Len(OnRoute) - 1)
                PathSoFar = OnRoute
                GoTo NewDire
            End If
        Else: GoTo NewDire
        End If
    Case 3
        Checked(a) = True
        If Movable(X + 1, Y) Then
            PathSoFar = PathSoFar & "r"
            If X + 1 = gX And Y = gY Then GoTo FoundRoute
            AIBoard(X, Y) = True
            OnRoute = OnRoute(X + 1, Y, gX, gY, PathSoFar)
            If OnRoute = PathSoFar Then 'No way found
                OnRoute = Left(OnRoute, Len(OnRoute) - 1)
                PathSoFar = OnRoute
                GoTo NewDire
            End If
        Else: GoTo NewDire
        End If
    Case 4
        Checked(a) = True
        If Movable(X, Y - 1) Then
            PathSoFar = PathSoFar & "u"
            If X = gX And Y - 1 = gY Then GoTo FoundRoute
            AIBoard(X, Y) = True
            OnRoute = OnRoute(X, Y - 1, gX, gY, PathSoFar)
            If OnRoute = PathSoFar Then 'No way found
                OnRoute = Left(OnRoute, Len(OnRoute) - 1)
                PathSoFar = OnRoute
                GoTo NewDire
            End If
        Else: GoTo NewDire
        End If
    End Select
    
    Stop 'it should NEVER get here
    Exit Function
FoundRoute:
    OnRoute = PathSoFar
    FinalPath = PathSoFar
End Function
Function Movable(X, Y) As Boolean
    Movable = True
    If X <= 0 Then Movable = False: Exit Function
    If X > BR Then Movable = False: Exit Function
    If Y <= 0 Then Movable = False: Exit Function
    If Y > HG Then Movable = False: Exit Function
    If Board(X, Y) Then Movable = False
    If AIBoard(X, Y) Then Movable = False
    'If Movable = False Then Stop
End Function

Public Sub MovePlayer() 'check if we can move to this square
    If P1.Path <> "" Then
        Select Case Mid(P1.Path, 1, 1)
        Case "l": P1.X = P1.X - 1
        Case "d": P1.Y = P1.Y + 1
        Case "r": P1.X = P1.X + 1
        Case "u": P1.Y = P1.Y - 1
        End Select
        P1.Path = Right(P1.Path, Len(P1.Path) - 1)
    End If
End Sub
