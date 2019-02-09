Public Class frmCalculator
    Public Shared intCount = 0
    Public Shared strOpsList As New List(Of String)
    Public Shared decNumList As New List(Of Decimal)
    Public Shared strOpsListTmp As New List(Of String)
    Public Shared decNumListTmp As New List(Of Decimal)


    Private Sub frmCalculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblDisplay.Text = ""
        lblEquation.Text = ""
    End Sub

    Private Sub btnNum1_Click(sender As Object, e As EventArgs) Handles btnNum1.Click
        lblDisplay.Text += "1"
    End Sub

    Private Sub btnNum2_Click(sender As Object, e As EventArgs) Handles btnNum2.Click
        lblDisplay.Text += "2"
    End Sub

    Private Sub btnNum3_Click(sender As Object, e As EventArgs) Handles btnNum3.Click
        lblDisplay.Text += "3"
    End Sub

    Private Sub btnNum4_Click(sender As Object, e As EventArgs) Handles btnNum4.Click
        lblDisplay.Text += "4"
    End Sub

    Private Sub btnNum5_Click(sender As Object, e As EventArgs) Handles btnNum5.Click
        lblDisplay.Text += "5"
    End Sub

    Private Sub btnNum6_Click(sender As Object, e As EventArgs) Handles btnNum6.Click
        lblDisplay.Text += "6"
    End Sub

    Private Sub btnNum7_Click(sender As Object, e As EventArgs) Handles btnNum7.Click
        lblDisplay.Text += "7"
    End Sub

    Private Sub btnNum8_Click(sender As Object, e As EventArgs) Handles btnNum8.Click
        lblDisplay.Text += "8"
    End Sub

    Private Sub btnNum9_Click(sender As Object, e As EventArgs) Handles btnNum9.Click
        lblDisplay.Text += "9"
    End Sub

    Private Sub btnNum0_Click(sender As Object, e As EventArgs) Handles btnNum0.Click
        lblDisplay.Text += "0"
    End Sub

    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If lblDisplay.Text = "" Then
            lblDisplay.Text += "0."
        Else
            If lblDisplay.Text.IndexOf(".") = -1 Then
                lblDisplay.Text += "."
            End If
        End If
    End Sub

    Private Sub btnSignChange_Click(sender As Object, e As EventArgs) Handles btnSignChange.Click
        If lblDisplay.Text = "" Then
            lblDisplay.Text += "-0"
        Else
            If lblDisplay.Text(0) <> "-" Then
                lblDisplay.Text = lblDisplay.Text.Insert(0, "-")
            Else
                lblDisplay.Text = lblDisplay.Text.Replace("-", "")
            End If
        End If
    End Sub

    Private Sub btnAllClear_Click(sender As Object, e As EventArgs) Handles btnAllClear.Click
        lblDisplay.Text = ""
        lblEquation.Text = ""
        decNumList.Clear()
        strOpsList.Clear()
    End Sub

    Private Sub btnBackspace_Click(sender As Object, e As EventArgs) Handles btnBackspace.Click
        If lblDisplay.Text <> "" Then
            lblDisplay.Text = lblDisplay.Text.Substring(0, lblDisplay.Text.Length - 1)
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lblDisplay.Text = ""
    End Sub

    Private Sub btnPlus_Click(sender As Object, e As EventArgs) Handles btnPlus.Click
        strOpsList.Add("+")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        strOpsList.Add("-")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        strOpsList.Add("*")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        strOpsList.Add("/")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnPow_Click(sender As Object, e As EventArgs) Handles btnPow.Click
        strOpsList.Add("^")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnPL_Click(sender As Object, e As EventArgs) Handles btnPL.Click
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
            strOpsList.Add("*")
        End If
        strOpsList.Add("(")
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnPR_Click(sender As Object, e As EventArgs) Handles btnPR.Click
        strOpsList.Add(")")
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If
        intCount += 1
        lblDisplay.Text = ""
        Display()
    End Sub

    Private Sub btnEqual_Click(sender As Object, e As EventArgs) Handles btnEqual.Click
        Dim intLevels As Integer = 0
        Dim intLeft As Integer
        Dim intRight As Integer
        Dim x As Integer = 0
        Dim blnContinuation As Boolean = True
        Dim blnLoop As Boolean

        'Add currently displayed number if present
        If lblDisplay.Text <> "" Then
            decNumList.Add(Convert.ToDecimal(lblDisplay.Text))
        End If

        lblEquation.Text += lblDisplay.Text & " ="

        blnLoop = strOpsList.IndexOf(")") <> -1

        While blnLoop
            'Get first ) location
            blnContinuation = True
            intLevels = 0
            While blnContinuation
                If strOpsList(x) = ")" Then
                    blnContinuation = False
                ElseIf strOpsList(x) = "(" Then
                    intLevels += 1
                    x += 1
                Else
                    x += 1
                End If
            End While

            intRight = x

            'MsgBox("Right at " & intRight)

            'Get the ( for )
            blnContinuation = True
            While blnContinuation
                If strOpsList(x) = "(" Then
                    blnContinuation = False
                End If
                x -= 1
            End While

            intLeft = x + 1

            'MsgBox("( at " & intLeft.ToString() & " ) at " & intRight.ToString())

            'Put ops and nums between parentheses into tmp lists
            For i = intLeft + 1 To intRight - 1
                strOpsListTmp.Add(strOpsList(i))
            Next

            For i = (intLeft - intLevels + 1) To (intRight - intLevels)
                decNumListTmp.Add(decNumList(i))
            Next
            'MsgBox(String.Join(",", decNumListTmp.ConvertAll(Of String)(Function(i As Integer) i.ToString())))

            'Calculate each
            Calculate("P", True)
            Calculate("MD", True)
            Calculate("AS", True)

            'Replace numbers between parentheses with results
            decNumList(intLeft - intLevels + 1) = decNumListTmp(0)
            decNumList.RemoveRange(intLeft - intLevels + 2, intRight - intLeft - 1)

            'Remove parentheses and ops between parentheses
            strOpsList.RemoveRange(intLeft, intRight - intLeft + 1)

            'Clear Tmp list
            decNumListTmp.Clear

            intLevels = 0
            x = 0

            'MsgBox(String.Join(",", decNumList.ConvertAll(Of String)(Function(i As Integer) i.ToString())))
            'MsgBox(String.Join(",", strOpsList))

            blnLoop = strOpsList.IndexOf(")") <> -1
        End While

        'MsgBox(String.Join(",", decNumList.ConvertAll(Of String)(Function(i As Integer) i.ToString())))

        'Commiting last operations 
        Calculate("P", False)
        Calculate("MD", False)
        Calculate("AS", False)

        lblDisplay.Text = decNumList(0).ToString()
        decNumList.Clear()
        strOpsList.Clear()
    End Sub

    Private Sub Operate(ByVal intIndex As Integer, ByRef chrOp As Char, ByRef blnPar As Boolean)
        Dim decSynthesis As Decimal
        'MsgBox("Operating at index " & intIndex)
        If blnPar Then
            Select Case chrOp
                Case "^"
                    decSynthesis = decNumListTmp(intIndex) ^ decNumListTmp(intIndex + 1)
                Case "*"
                    decSynthesis = decNumListTmp(intIndex) * decNumListTmp(intIndex + 1)
                Case "/"
                    decSynthesis = decNumListTmp(intIndex) / decNumListTmp(intIndex + 1)
                Case "+"
                    decSynthesis = decNumListTmp(intIndex) + decNumListTmp(intIndex + 1)
                Case "-"
                    decSynthesis = decNumListTmp(intIndex) - decNumListTmp(intIndex + 1)
            End Select
            decNumListTmp(intIndex) = decSynthesis
            'Removing numbers and operator
            decNumListTmp.RemoveAt(intIndex + 1)
            strOpsListTmp.RemoveAt(intIndex)
        Else
            Select Case chrOp
                Case "^"
                    decSynthesis = decNumList(intIndex) ^ decNumList(intIndex + 1)
                Case "*"
                    decSynthesis = decNumList(intIndex) * decNumList(intIndex + 1)
                Case "/"
                    Try
                        decSynthesis = decNumList(intIndex) / decNumList(intIndex + 1)
                    Catch ex As DivideByZeroException
                        MsgBox("Equation has dividing by zero.",, "Error: Divide by Zero")
                        Return
                    End Try
                Case "+"
                    decSynthesis = decNumList(intIndex) + decNumList(intIndex + 1)
                Case "-"
                    decSynthesis = decNumList(intIndex) - decNumList(intIndex + 1)
            End Select
            decNumList(intIndex) = decSynthesis
            'Removing numbers and operator
            decNumList.RemoveAt(intIndex + 1)
            strOpsList.RemoveAt(intIndex)
        End If
    End Sub

    'Functions for use
    Private Sub Calculate(ByRef chrOp As Char, ByRef blnPar As Boolean)
        'Commits all of the certain ops in equation
        'Dim ops As String = String.Join(",", strOpsList)
        Dim intIndexM As Integer
        Dim intIndexD As Integer
        Dim intIndexA As Integer
        Dim intIndexS As Integer
        Dim intIndexP As Integer
        'MsgBox(ops)
        If blnPar Then
            Select Case chrOp
                Case "P"
                    intIndexP = strOpsListTmp.IndexOf("^")
                    While intIndexP <> -1
                        Operate(intIndexP, "^", True)
                        intIndexP = strOpsListTmp.IndexOf("^")
                    End While
                Case "MD"
                    intIndexD = strOpsListTmp.IndexOf("/")
                    intIndexM = strOpsListTmp.IndexOf("*")
                    'If both * and / are present, do the operations that come first
                    While intIndexD <> -1 And intIndexM <> -1
                        If intIndexD < intIndexM Then
                            'Do Division
                            Operate(intIndexD, "/", True)
                        ElseIf intIndexD > intIndexM Then
                            'Do Multiplication
                            Operate(intIndexM, "*", True)
                        Else
                            MsgBox("Error: * and / sign somehow in same place")
                        End If
                        intIndexD = strOpsListTmp.IndexOf("/")
                        intIndexM = strOpsListTmp.IndexOf("*")
                    End While
                    While intIndexD <> -1
                        'Do Division
                        Operate(intIndexD, "/", True)
                        intIndexD = strOpsListTmp.IndexOf("/")
                    End While
                    While intIndexM <> -1
                        'Do Division
                        Operate(intIndexM, "*", True)
                        intIndexM = strOpsListTmp.IndexOf("*")
                    End While
                Case "AS"
                    intIndexA = strOpsListTmp.IndexOf("+")
                    intIndexS = strOpsListTmp.IndexOf("-")
                    While intIndexA <> -1 And intIndexS <> -1
                        If intIndexA < intIndexS Then
                            'Do Addition
                            Operate(intIndexA, "+", True)
                        ElseIf intIndexA > intIndexS Then
                            'Do Multiplication
                            Operate(intIndexS, "-", True)
                        Else
                            MsgBox("Error: + and - sign somehow in same place")
                        End If
                        intIndexA = strOpsListTmp.IndexOf("+")
                        intIndexS = strOpsListTmp.IndexOf("-")
                    End While
                    'Do the final add or subtracts
                    While intIndexS <> -1
                        'Do Division
                        Operate(intIndexS, "-", True)
                        intIndexS = strOpsListTmp.IndexOf("-")
                    End While
                    While intIndexA <> -1
                        'Do Division
                        Operate(intIndexA, "+", True)
                        intIndexA = strOpsListTmp.IndexOf("+")
                    End While
            End Select
        Else
            Select Case chrOp
                Case "P"
                    intIndexP = strOpsList.IndexOf("^")
                    While intIndexP <> -1
                        Operate(intIndexP, "^", False)
                        intIndexP = strOpsList.IndexOf("^")
                    End While
                Case "MD"
                    intIndexD = strOpsList.IndexOf("/")
                    intIndexM = strOpsList.IndexOf("*")
                    'If both * and / are present, do the operations that come first
                    While intIndexD <> -1 And intIndexM <> -1
                        If intIndexD < intIndexM Then
                            'Do Division
                            Operate(intIndexD, "/", False)
                        ElseIf intIndexD > intIndexM Then
                            'Do Multiplication
                            Operate(intIndexM, "*", False)
                        Else
                            MsgBox("Error: * and / sign somehow in same place")
                        End If
                        intIndexD = strOpsList.IndexOf("/")
                        intIndexM = strOpsList.IndexOf("*")
                    End While
                    While intIndexD <> -1
                        'Do Division
                        Operate(intIndexD, "/", False)
                        intIndexD = strOpsList.IndexOf("/")
                    End While
                    While intIndexM <> -1
                        'Do Division
                        Operate(intIndexM, "*", False)
                        intIndexM = strOpsList.IndexOf("*")
                    End While
                Case "AS"
                    intIndexA = strOpsList.IndexOf("+")
                    intIndexS = strOpsList.IndexOf("-")
                    While intIndexA <> -1 And intIndexS <> -1
                        If intIndexA < intIndexS Then
                            'Do Addition
                            Operate(intIndexA, "+", False)
                        ElseIf intIndexA > intIndexS Then
                            'Do Multiplication
                            Operate(intIndexS, "-", False)
                        Else
                            MsgBox("Error: + and - sign somehow in same place")
                        End If
                        intIndexA = strOpsList.IndexOf("+")
                        intIndexS = strOpsList.IndexOf("-")
                    End While
                    'Do the final add or subtracts
                    While intIndexS <> -1
                        'Do Division
                        Operate(intIndexS, "-", False)
                        intIndexS = strOpsList.IndexOf("-")
                    End While
                    While intIndexA <> -1
                        'Do Division
                        Operate(intIndexA, "+", False)
                        intIndexA = strOpsList.IndexOf("+")
                    End While
            End Select
        End If

    End Sub

    Private Sub Display()
        'Displays the equation
        Dim x As Integer = 0
        lblEquation.Text = ""
        'For i = 0 To (strOpsList.Count - 1)
        '    If strOpsList(i) = "(" Then
        '        lblEquation.Text += strOpsList(i)
        '        lblEquation.Text += " "
        '        x += 1
        '    ElseIf strOpsList(i) = ")" Then
        '        lblEquation.Text += decNumList(i - x).ToString()
        '        lblEquation.Text += strOpsList(i)
        '        lblEquation.Text += " "
        '        x -= 1
        '    ElseIf strOpsList(i).IndexOfAny("^+-*/") <> -1 Then
        '        If i > 0 Then
        '            If strOpsList(i - 1) <> ")" Then
        '                Try
        '                    lblEquation.Text += decNumList(i - x).ToString()
        '                Catch ex As ArgumentException
        '                    MsgBox("Please input correctly",, "Error: Invalid Input")
        '                    Return
        '                End Try
        '            End If
        '        ElseIf i = 0 Then
        '            lblEquation.Text += decNumList(i).ToString()
        '        End If
        '        lblEquation.Text += " "
        '        lblEquation.Text += strOpsList(i)
        '        lblEquation.Text += " "
        '    End If
        'Next
    End Sub

    Private Sub btnPi_Click(sender As Object, e As EventArgs) Handles btnPi.Click
        'Displays Pi
        Dim Pi As Decimal = 3.1415926535897931
        lblDisplay.Text = Pi.ToString()
    End Sub

    Private Sub btnFraction_Click(sender As Object, e As EventArgs) Handles btnFraction.Click
        'Gets and displayx 1/x; x is currently displayed number
        Dim Synthesis As Decimal

        If lblDisplay.Text <> "" Then
            Synthesis = 1 / Convert.ToDecimal(lblDisplay.Text)
            lblDisplay.Text = Synthesis.ToString()
        Else
            MsgBox("Enter a number")
        End If
    End Sub

    Private Sub btnSquare_Click(sender As Object, e As EventArgs) Handles btnSquare.Click
        'Gets and displays the squareroot of x; x is currently displayed number
        Dim decInput As Decimal
        Dim Synthesis As Decimal

        If lblDisplay.Text <> "" Then
            decInput = Convert.ToDecimal(lblDisplay.Text)
            If decInput > 0 Then
                Synthesis = Math.Sqrt(decInput)
                lblDisplay.Text = Synthesis.ToString()
            Else
                MsgBox("Non-real Answer")
            End If
        Else
            MsgBox("Enter a number")
        End If
    End Sub

    Private Sub btnE_Click(sender As Object, e As EventArgs) Handles btnE.Click
        'Display E
        Dim Euler As Decimal = 2.7182818284590451
        lblDisplay.Text = Euler.ToString()
    End Sub

    Private Sub btnFactorial_Click(sender As Object, e As EventArgs) Handles btnFactorial.Click
        'Get and display the factorial for the displayed number
        Dim Synthesis As Decimal = 1
        Dim decInput As Decimal

        If lblDisplay.Text <> "" Then
            decInput = Convert.ToDecimal(lblDisplay.Text)
            If decInput = 0 Then
                Synthesis = 1
                lblDisplay.Text = Synthesis.ToString()
            ElseIf decInput < 0 Then
                MsgBox("Can't congfigure negative number.")
                Return
            Else
                MsgBox("This is > 1")
                If Math.Floor(decInput) = decInput Then
                    For i = 1 To decInput
                        Synthesis = Synthesis * i
                    Next
                    lblDisplay.Text = Synthesis.ToString()
                Else
                    MsgBox("Can't congfigure decimal number.")
                    Return
                End If
            End If
        Else
            MsgBox("Enter a number")
        End If
    End Sub
End Class
