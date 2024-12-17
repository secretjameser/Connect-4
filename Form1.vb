Imports System.Net.Sockets

Public Class Connect4
    Dim Rows As Integer = 6
    Dim Columns As Integer = 7
    Dim Board(Rows - 1, Columns - 1) As (Boolean, Boolean)
    Dim Table() As (Integer, Integer) 'later
    Dim CurrentPlayer As Integer = 1

    Private Sub Connect4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Randomize()

        For Each control As Control In Me.Controls
            If TypeOf control Is PictureBox Then
                AddHandler control.MouseDown, AddressOf Cell_Click
            End If
        Next

        Call InitializeBoard()
    End Sub

    Private Sub BtnNew_Click(sender As Object, e As EventArgs) Handles BtnNew.Click
        Call InitializeBoard()
    End Sub

    Private Sub BtnDone_Click(sender As Object, e As EventArgs) Handles BtnDone.Click
        End
    End Sub

    Sub InitializeBoard()
        CurrentPlayer = 1

        lblTurn.Text = "Red's Turn"

        For row As Integer = 0 To Rows - 1
            For col As Integer = 0 To Columns - 1
                Board(row, col) = (False, False)
            Next
        Next

        Call DrawBoard()
    End Sub

    Sub DrawBoard() 'CHANGE?
        Dim cell As PictureBox
        Dim row, col As Integer
        For Each control As Control In Me.Controls
            If TypeOf control Is PictureBox Then
                cell = DirectCast(control, PictureBox)
                row = cell.Top \ 64
                col = cell.Left \ 64

                If Board(row, col).Item1 Then
                    cell.Image = Image.FromFile("red.png")
                ElseIf Board(row, col).item2 Then
                    cell.Image = Image.FromFile("yellow.png")
                Else
                    cell.Image = Image.FromFile("white.png")
                End If
            End If
        Next
    End Sub

    Sub Cell_Click(sender As Object, e As EventArgs)
        Dim clickedCell As PictureBox = DirectCast(sender, PictureBox)
        Dim column As Integer = clickedCell.Left \ 64

        If DropPiece(column, True, Board) Then
            Call DrawBoard()

            If CheckWin(Board, True) Then
                MsgBox("Red wins!")
                Call InitializeBoard()
                Exit Sub
            Else
                'If CurrentPlayer = 1 Then
                '    CurrentPlayer = 2

                '    lblTurn.Text = "Yellow's Turn"
                'Else
                '    CurrentPlayer = 1

                '    lblTurn.Text = "Red's Turn"
                'End If
            End If

            If chkComp.Checked Then
                Call DrawBoard()
                Call Comp()
                Call DrawBoard()

                If CheckWin(Board, False) Then
                    lblTurn.Text = "Yellow's Turn"
                    MsgBox("Yellow wins!")
                    Call InitializeBoard()
                    Exit Sub
                End If

            Else
                If CurrentPlayer = 1 Then
                    CurrentPlayer = 2

                    lblTurn.Text = "Yellow's Turn"
                Else
                    CurrentPlayer = 1

                    lblTurn.Text = "Red's Turn"
                End If
            End If
        End If
    End Sub

    Sub Comp() 'NEEDS UPDATE
        Dim x, minScore, row, score, depth As Integer
        x = -1
        minScore = Integer.MaxValue
        depth = Val(txtDepth.Text)

        For col As Integer = 0 To Columns - 1
            If Board(0, col) = 0 Then
                row = GetDropRow(col, Board)
                Board(row, col) = 2
                score = Alphabeta(Board, depth, Integer.MinValue, Integer.MaxValue, True)
                Board(row, col) = 0

                If score < minScore Then
                    minScore = score
                    x = col
                End If
            End If
        Next

        If x >= 0 Then
            DropPiece(x, False, Board)
        End If
    End Sub

    Function Alphabeta(node(,) As Integer, depth As Integer, alpha As Integer, beta As Integer, maximizingPlayer As Boolean) As Integer
        Dim value, row As Integer

        If depth = 0 Then
            Return EvaluateBoard(node)
        End If

        If CheckWin(node, True) Then
            Return Integer.MaxValue - (100 - depth)
        End If

        If CheckWin(node, False) Then
            Return Integer.MinValue + (100 - depth)
        End If

        If CheckTie(node) Then
            Return 0
        End If

        If maximizingPlayer Then
            value = Integer.MinValue

            For col As Integer = 0 To Columns - 1
                If node(0, col) = 0 Then
                    row = GetDropRow(col, node)

                    node(row, col) = 1

                    value = Math.Max(value, Alphabeta(node, depth - 1, alpha, beta, False))

                    node(row, col) = 0

                    alpha = Math.Max(alpha, value)

                    If value >= beta Then
                        Exit For
                    End If
                End If
            Next

            Return value
        Else
            value = Integer.MaxValue

            For col As Integer = 0 To Columns - 1
                If node(0, col) = 0 Then
                    row = GetDropRow(col, node)

                    node(row, col) = 2

                    value = Math.Min(value, Alphabeta(node, depth - 1, alpha, beta, True))

                    node(row, col) = 0

                    beta = Math.Min(beta, value)

                    If value <= alpha Then
                        Exit For
                    End If
                End If
            Next

            Return value
        End If
    End Function

    Function GetDropRow(column As Integer, node(,) As Integer) As Integer
        For row As Integer = Rows - 1 To 0 Step -1
            If node(row, column) = 0 Then
                Return row
            End If
        Next
        Return -1
    End Function

    Function EvaluateBoard(node(,) As Integer) As Integer
        Dim score As Integer = 0
        For row As Integer = 0 To Rows - 1
            For col As Integer = 0 To Columns - 1
                If node(row, col) <> 0 Then
                    score += EvaluatePosition(row, col, node)
                End If
            Next
        Next
        Return score
    End Function

    Function EvaluatePosition(row As Integer, col As Integer, node(,) As Integer) As Integer
        Dim score, count, r, c As Integer
        Dim directions As (Integer, Integer)() = {(1, 0), (1, 1), (0, 1), (-1, 1), (-1, 0), (-1, -1), (0, -1), (1, -1)}

        For Each dire In directions
            count = 0
            For j As Integer = 0 To 3
                r = row + j * dire.Item1
                c = col + j * dire.Item2

                If r >= 0 AndAlso r < Rows AndAlso c >= 0 AndAlso c < Columns Then
                    If node(r, c) = 1 Then
                        count += 1
                    ElseIf node(r, c) = 2 Then
                        count = 0
                        Exit For
                    ElseIf node(r, c) = 0 Then
                        Exit For
                    End If
                Else
                    count = 0
                    Exit For
                End If
            Next

            score += count ^ 8 'idk

            count = 0
            For j As Integer = 0 To 3
                r = row + j * dire.Item1
                c = col + j * dire.Item2

                If r >= 0 AndAlso r < Rows AndAlso c >= 0 AndAlso c < Columns Then
                    If node(r, c) = 2 Then
                        count += 1
                    ElseIf node(r, c) = 1 Then
                        count = 0
                        Exit For
                    End If
                Else
                    count = 0
                    Exit For
                End If
            Next

            score -= count ^ 8 'idk
        Next

        Return score
    End Function

    Function DropPiece(column As Integer, maximizingPlayer As Boolean, node(,) As (Boolean, Boolean)) As Boolean
        For row As Integer = Rows - 1 To 0 Step -1
            If Not node(row, column).Item1 And Not node(row, column).Item2 Then 'oop
                If maximizingPlayer Then
                    node(row, column).Item1 = True
                Else
                    node(row, column).Item2 = True
                End If

                Return True
            End If
        Next

        Return False
    End Function

    Function CheckTie(node(,) As Integer) As Boolean
        For col As Integer = 0 To Columns - 1
            If node(0, col) = 0 Then
                Return False
            End If
        Next
        Return True
    End Function

    Function CheckWin(node(,) As (Boolean, Boolean), maximizingPlayer As Boolean) As Boolean
        For row As Integer = 0 To Rows - 1
            For col As Integer = 0 To Columns - 1
                If Not node(row, col).Item1 And Not node(row, col).Item2 Then 'oop
                    If CheckDirection(row, col, 1, 0, node, maximizingPlayer) Then
                        Return True
                    End If

                    If CheckDirection(row, col, 0, 1, node, maximizingPlayer) Then
                        Return True
                    End If

                    If CheckDirection(row, col, 1, 1, node, maximizingPlayer) Then
                        Return True
                    End If

                    If CheckDirection(row, col, 1, -1, node, maximizingPlayer) Then
                        Return True
                    End If
                End If
            Next
        Next
        Return False
    End Function

    Function CheckDirection(row As Integer, col As Integer, dRow As Integer, dCol As Integer, node(,) As Integer, maximizingPlayer As Boolean) As Boolean
        Dim count As Integer = 0
        Dim player As Integer = If(maximizingPlayer = True, 1, 2)
        Dim r, c As Integer

        For i As Integer = 0 To 3
            r = row + i * dRow
            c = col + i * dCol

            If r >= 0 AndAlso r < Rows AndAlso c >= 0 AndAlso c < Columns AndAlso node(r, c) = player Then
                count += 1
            Else
                Exit For
            End If
        Next

        Return count = 4
    End Function
End Class
