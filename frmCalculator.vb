Public Class frmCalculator
    Public Shared decTmp As Decimal
    Public Shared strFuncs() As String = {"add", "minus", "multiply", "divide", "power"}
    Public Shared strChoice As String = ""

    'Functions for changing display
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

    Private Sub btnSign_Click(sender As Object, e As EventArgs) Handles btnSign.Click
        If lblDisplay.Text = "" Then
            lblDisplay.Text += "-"
        Else
            If lblDisplay.Text(0) <> "-" Then
                lblDisplay.Text = lblDisplay.Text.Insert(0, "-")
            Else
                lblDisplay.Text = lblDisplay.Text.Replace("-", "")
            End If
        End If
    End Sub

    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If lblDisplay.Text.IndexOf(".") = -1 Then
            lblDisplay.Text += "."
        End If
    End Sub

    Private Sub btnE_Click(sender As Object, e As EventArgs) Handles btnE.Click
        lblDisplay.Text = Math.E.ToString()
    End Sub

    Private Sub btnPi_Click(sender As Object, e As EventArgs) Handles btnPi.Click
        lblDisplay.Text = Math.PI.ToString()
    End Sub
    'Mathematical Functions
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ToRegister()
        strChoice = "add"
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        ToRegister()
        strChoice = "minus"
    End Sub

    Private Sub btnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        ToRegister()
        strChoice = "multiply"
    End Sub

    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        ToRegister()
        strChoice = "divide"
    End Sub

    Private Sub btnPower_Click(sender As Object, e As EventArgs) Handles btnPower.Click
        ToRegister()
        strChoice = "power"
    End Sub

    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        Dim decCurrent As Decimal

        'Get current number
        Try
            decCurrent = Convert.ToDecimal(lblDisplay.Text)
        Catch ex As FormatException
            MsgBox("Please enter a number.")
            Return
        End Try

        'Call Compute Function
        lblDisplay.Text = Calculate(decTmp, decCurrent, Array.IndexOf(strFuncs, strChoice))
        strChoice = ""
        decTmp = Nothing
    End Sub

    'Functions for Running Calculator
    Private Sub ToRegister()
        'Adds currently displayed number to register
        decTmp = Convert.ToDecimal(lblDisplay.Text)

        lblDisplay.Text = ""
    End Sub

    Private Function Calculate(decFirst As Decimal, decSecond As Decimal, intFunc As Integer)
        'Calculates
        'Needs first number, second number, and integer code
        Dim decSynthesis As String = ""

        Select Case intFunc
            Case -1
                decSynthesis = decSecond
            Case 0
                decSynthesis = decFirst + decSecond
            Case 1
                decSynthesis = decFirst - decSecond
            Case 2
                decSynthesis = decFirst * decSecond
            Case 3
                Try
                    decSynthesis = decFirst / decSecond
                Catch ex As DivideByZeroException
                    MsgBox("Cannot divide by 0",, "Error: Divide by 0")
                    Return ""
                End Try
            Case 4
                decSynthesis = decFirst ^ decSecond
        End Select

        Return decSynthesis.ToString()
    End Function

    'Clearing Functions
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lblDisplay.Text = ""
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        lblDisplay.Text = ""
        decTmp = 0
    End Sub

    Private Sub btnClearLast_Click(sender As Object, e As EventArgs) Handles btnClearLast.Click
        If lblDisplay.Text <> "" Then
            lblDisplay.Text = lblDisplay.Text.Substring(0, lblDisplay.Text.Length - 1)
        End If
    End Sub

    'Special Functions; Functions that do not require the equal sign to be computed
    Private Sub btnFraction_Click(sender As Object, e As EventArgs) Handles btnFraction.Click
        If lblDisplay.Text = "" Or lblDisplay.Text = "-" Then
            MsgBox("Please input a value.",, "Error: No value")
            Return
        End If

        'Commence to calculate 1/x, where x is the number on display
        Dim decInput As Decimal = Convert.ToDecimal(lblDisplay.Text)
        Dim decSynthesis As Decimal

        Try
            decSynthesis = 1 / decInput
        Catch ex As DivideByZeroException
            MsgBox("Cannot divide by 0",, "Error: Divide by 0")
            Return
        End Try

        lblDisplay.Text = decSynthesis.ToString()
    End Sub

    Private Sub btnSquareRoot_Click(sender As Object, e As EventArgs) Handles btnSquareRoot.Click
        'Get Square root of currnet number
        If lblDisplay.Text = "" Or lblDisplay.Text = "-" Then
            MsgBox("Please input a value.",, "Error: No value")
            Return
        End If

        Dim decInput As Decimal = Convert.ToDecimal(lblDisplay.Text)
        Dim decSynthesis As Decimal

        Try
            decSynthesis = Math.Sqrt(decInput)
        Catch ex As OverflowException
            MsgBox("No real square root",, "Error: Imaginary Root")
            Return
        End Try
        lblDisplay.Text = decSynthesis.ToString()
    End Sub
End Class
