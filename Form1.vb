Public Class Form1
    Dim angka1 As Double
    Dim [operator] As String
    Dim baru As Boolean = True
    Dim temp As Double
    Dim angka2 As Double

    Private Sub B0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B0.Click, B1.Click, B2.Click, B3.Click, B4.Click, B5.Click, B6.Click, B7.Click, B8.Click, B9.Click
        TextBox1.Text = TextBox1.Text & sender.text
    End Sub

    Private Sub Bkali_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bkali.Click
        If baru = True Then
            angka1 = Val(TextBox1.Text)
            TextBox1.Text = ""
            TextBox1.Focus()
            [operator] = "x"
        ElseIf baru = False Then
            If [operator] = "+" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "x"
            ElseIf [operator] = "-" Then
                temp = angka1 - Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "x"
            ElseIf [operator] = ":" Then
                temp = angka1 / Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "x"
            ElseIf [operator] = "x" Then
                temp = angka1 * Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "x"
            End If
        End If

        baru = False

    End Sub

    Private Sub Btambah_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btambah.Click

        If baru = True Then
            angka1 = Val(TextBox1.Text)
            TextBox1.Text = ""
            TextBox1.Focus()
            [operator] = "+"
        ElseIf baru = False Then
            If [operator] = "+" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "+"
            ElseIf [operator] = "-" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "+"
            ElseIf [operator] = "x" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "+"
            ElseIf [operator] = ":" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "+"
            End If
        End If
        baru = True

    End Sub

    Private Sub Bkurang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bkurang.Click

        If baru = True Then
            angka1 = Val(TextBox1.Text)
            TextBox1.Text = ""
            TextBox1.Focus()
            [operator] = "-"
        ElseIf baru = False Then
            If [operator] = "+" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "-"
            ElseIf [operator] = "-" Then
                temp = angka1 - Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "-"
            ElseIf [operator] = "x" Then
                temp = angka1 * Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "-"
            ElseIf [operator] = ":" Then
                temp = angka1 / Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "-"

            End If
            End If
            baru = False

    End Sub

    Private Sub Bbagi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bbagi.Click

        If baru = True Then
            angka1 = Val(TextBox1.Text)
            TextBox1.Text = ""
            TextBox1.Focus()
            [operator] = "/"
        ElseIf baru = False Then
            If [operator] = "/" Then
                temp = angka1 / Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "/"
            ElseIf [operator] = "x" Then
                temp = angka1 * Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "/"
            ElseIf [operator] = "+" Then
                temp = angka1 + Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "/"
            ElseIf [operator] = "-" Then
                temp = angka1 - Val(TextBox1.Text)
                angka1 = temp
                TextBox1.Text = ""
                [operator] = "/"

            End If

        End If
        baru = False

    End Sub
    Dim samadengan As Double

    Private Sub Bsamadengan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bsamadengan.Click
        angka2 = Val(TextBox1.Text)
        Select Case [operator]
            Case "+"
                samadengan = angka1 + angka2
                TextBox1.Text = samadengan.ToString()
                baru = True
            Case "-"
                samadengan = angka1 - angka2
                TextBox1.Text = samadengan.ToString
                baru = True
            Case "/"
                samadengan = angka1 / angka2
                TextBox1.Text = samadengan.ToString
                baru = True
            Case "x"
                samadengan = angka1 * angka2
                TextBox1.Text = samadengan.ToString
        End Select
        TextBox1.Text = samadengan.ToString

    End Sub

    Private Sub Bhapus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bhapus.Click
        TextBox1.Text = ""
        angka1 = 0
        temp = 0
        baru = True

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Btitik_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btitik.Click

    End Sub
    Dim sapiObject
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sapiObject = CreateObject("SAPI.spvoice")

        sapiObject.speak(TextBox1.Text)

    End Sub
End Class






