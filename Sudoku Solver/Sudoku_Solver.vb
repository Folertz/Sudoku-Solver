Imports System.Threading

Public Class SudokuSolver

#Region "Global Declarations"

    Dim cell(0 To 8, 0 To 8) As SudokuNumBoxes
    Dim grid(0 To 8, 0 To 8) As String
    Dim BackTrackOn As Boolean = False
    Dim InstNumber As Integer

#End Region

#Region "Grid Drawing and Spacing"

    Private Sub SudokuSolver_Load(sender As Object, e As EventArgs) Handles Me.Load
        UpdateInfo(" - Application Started")
        Dim XSpace As Integer
        Dim YSpace As Integer
        For x As Integer = 0 To 8
            For y As Integer = 0 To 8
                cell(x, y) = New SudokuNumBoxes
                cell(x, y).Text = ""
                cell(x, y).Width = 20
                cell(x, y).Height = 20
                cell(x, y).MaxLength = 1
                cell(x, y).Font = New Font("Arial", 8.25, FontStyle.Regular)
                cell(x, y).TextAlign = HorizontalAlignment.Center
                XSpace = 0
                YSpace = 0
                If x > 2 Then
                    XSpace = 4
                End If
                If x > 5 Then
                    XSpace = 8
                End If
                If y > 2 Then
                    YSpace = 4
                End If
                If y > 5 Then
                    YSpace = 8
                End If
                cell(x, y).Location = New Point(12 + x * 20 + XSpace, 12 + y * 20 + YSpace)
                Me.Controls.Add(cell(x, y))
                AddHandler cell(x, y).TextChanged, AddressOf cell_changed
            Next
        Next
    End Sub

#End Region

#Region "Error Proofing"

    Private Sub cell_changed(sender As Object, e As EventArgs)
        If BackTrackOn Then Return
        For x As Integer = 0 To 8
            For y As Integer = 0 To 8
                grid(x, y) = cell(x, y).Text
                cell(x, y).ForeColor = Color.Black
                ButtonSolve.Enabled = True
                UpdateStatus("Can be solved..")
            Next
        Next
        For x As Integer = 0 To 8
            For y As Integer = 0 To 8
                If check_rows(x, y) Then
                    If check_columns(x, y) Then
                        If Not check_boxes(x, y) Then
                            cell(x, y).ForeColor = Color.Red
                            ButtonSolve.Enabled = False
                            UpdateStatus("Error : Matching numbers found in a ""3 x 3"" block..")
                        End If
                    Else
                        cell(x, y).ForeColor = Color.Red
                        ButtonSolve.Enabled = False
                        UpdateStatus("Error : Matching numbers found in a column..")
                    End If
                Else
                    cell(x, y).ForeColor = Color.Red
                    ButtonSolve.Enabled = False
                    UpdateStatus("Error : Matching numbers found in a row..")
                End If
            Next
        Next
    End Sub

    Function check_rows(ByVal Xsend As Integer, ByVal YSend As Integer) As Boolean
        Dim noclash As Boolean = True
        For x As Integer = 0 To 8
            If grid(x, CInt(YSend)) <> "" Then
                If x <> Xsend Then
                    If grid(x, CInt(YSend)) = grid(CInt(Xsend), CInt(YSend)) Then
                        noclash = False
                    End If
                End If
            End If
        Next
        Return noclash
    End Function

    Function check_columns(ByVal Xsend As Integer, ByVal YSend As Integer) As Boolean
        Dim noclash As Boolean = True
        For y As Integer = 0 To 8
            If grid(CInt(Xsend), y) <> "" Then
                If y <> YSend Then
                    If grid(CInt(Xsend), y) = grid(CInt(Xsend), CInt(YSend)) Then
                        noclash = False
                    End If
                End If
            End If
        Next
        Return noclash
    End Function

    Function check_boxes(ByVal Xsend As Integer, ByVal YSend As Integer) As Boolean
        Dim noclash As Boolean = True
        Dim XStart As Integer
        Dim YStart As Integer
        If Xsend < 3 Then
            XStart = 0
        ElseIf Xsend < 6 Then
            XStart = 3
        Else
            XStart = 6
        End If
        If YSend < 3 Then
            YStart = 0
        ElseIf YSend < 6 Then
            YStart = 3
        Else
            YStart = 6
        End If
        For y As Integer = YStart To (YStart + 2)
            For x As Integer = XStart To (XStart + 2)
                If grid(x, y) <> "" Then
                    If Not (x = Xsend And y = YSend) Then
                        If grid(x, y) = grid(CInt(Xsend), CInt(YSend)) Then
                            noclash = False
                        End If
                    End If
                End If
            Next
        Next
        Return noclash
    End Function

#End Region

#Region "Solver"

    Private Delegate Sub _BackTrackThread()

    Private Sub ButtonSolve_Click(sender As Object, e As EventArgs) Handles ButtonSolve.Click
        Dim T As Thread = New Thread(AddressOf DoWork)
        With T
            .Name = "Solving Thread"
            .Priority = ThreadPriority.BelowNormal
            .IsBackground = True
            .Start()
        End With
        InstNumber = 0
    End Sub

    Private Sub DoWork(ByVal BackTrackOn As Boolean)
        UpdateInfo(" - Thread for solving started.")
        UpdateStatus("Solving..")
        BackTrackThread()
    End Sub

    Sub BackTrackThread()
        Thread.Sleep(1)
        If (InvokeRequired) Then
            Invoke(New _BackTrackThread(AddressOf BackTrackThread))
        Else
            BackTrackOn = True
            For x As Integer = 0 To 8
                For y As Integer = 0 To 8
                    grid(x, y) = cell(x, y).Text
                Next
            Next
            BruteForceBacktracking(0, 0)
            For x As Integer = 0 To 8
                For y As Integer = 0 To 8
                    cell(x, y).Text = grid(x, y)
                Next
            Next
            UpdateStatus("Successfully complete the puzzle.")
            UpdateInfo(String.Format(" - Completed puzzle after {0} moves.", InstNumber.ToString))
            BackTrackOn = False
        End If
    End Sub

    Function BruteForceBacktracking(ByVal x As Integer, ByVal y As Integer) As Boolean
        Dim Number As Integer = 1
        If grid(x, y) = "" Then
            Do
                grid(x, y) = CStr(Number)
                If check_rows(x, y) Then
                    If check_columns(x, y) Then
                        If check_boxes(x, y) Then
                            InstNumber = InstNumber + 1
                            y = y + 1
                            If y = 9 Then
                                y = 0
                                x = x + 1
                                If x = 9 Then Return True
                            End If
                            If BruteForceBacktracking(x, y) Then Return True
                            y = y - 1
                            If y < 0 Then
                                y = 8
                                x = x - 1
                            End If
                        End If
                    End If
                End If
                Number = Number + 1

            Loop Until Number = 10

            grid(x, y) = ""
            Return False
        Else
            y = y + 1
            If y = 9 Then
                y = 0
                x = x + 1

                If x = 9 Then Return True
            End If
            Return BruteForceBacktracking(x, y)
        End If

    End Function

    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        For x As Integer = 0 To 8
            For y As Integer = 0 To 8
                cell(x, y).Text = ""
            Next
        Next
        UpdateInfo(" - Grid was cleared.")
    End Sub

    Private Sub ClearLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearLogToolStripMenuItem.Click
        EventLog.Clear()
    End Sub

#End Region

#Region "Classes"
    Class SudokuNumBoxes
        Inherits TextBox
        Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
            If Char.IsDigit(e.KeyChar) Or e.KeyChar = " " Or e.KeyChar = ControlChars.Back Then
                e.Handled = False
            Else
                e.Handled = True
            End If
            If e.KeyChar = " " Or e.KeyChar = "0" Then
                e.KeyChar = ControlChars.Back
            End If
        End Sub
    End Class
#End Region

#Region "UpdateLog Subs"

    Private Delegate Sub _UpdateInfo(ByVal LogInfo As String)
    Private Sub UpdateInfo(ByVal LogInfo As String)
        If (InvokeRequired) Then
            Invoke(New _UpdateInfo(AddressOf UpdateInfo), LogInfo)
        Else
            With EventLog
                .AppendText(TimeOfDay.ToLongTimeString + LogInfo)
                .AppendText(Environment.NewLine)
            End With
        End If
    End Sub
    Private Delegate Sub _UpdateStatus(ByVal StatusInfo As String)
    Private Sub UpdateStatus(ByVal StatusInfo As String)
        If (InvokeRequired) Then
            Invoke(New _UpdateStatus(AddressOf UpdateStatus), StatusInfo)
        Else
            StatusError.Text = StatusInfo
        End If
    End Sub


#End Region

End Class


