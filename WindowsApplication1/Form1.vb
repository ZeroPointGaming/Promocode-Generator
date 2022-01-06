Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.Windows.Controls

Public Class Form1
    Private randomBytes() As Byte
    Private randomInt32Value As Integer
    Private possibleChars As String
    Private len As Int32
    Private GetRandomInt32Value As New RandomInt32Value
    Private GetPasswordGenProfiler As New PasswordGenProfiler

    Public ReadOnly Property LineCount As Integer

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        possibleChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        len = "15"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim cpossibleChars() As Char
            cpossibleChars = possibleChars.ToCharArray()
            If cpossibleChars.Length < 1 Then
                MessageBox.Show("You must enter one or more possible characters.")
                Return
            End If
            If len < 4 Then
                MessageBox.Show(String.Format("Please choose a password length. That length must be a value between {0} and {1}. Note: values above 1,000 might take a LONG TIME to process on some computers.", 4, Int32.MaxValue))
                Return
            End If

            Dim builder As New StringBuilder()

            For i As Integer = 0 To len - 1
                Dim randInt32 As Integer = GetRandomInt32Value.GetRandomInt()
                Dim r As New Random(randInt32)

                Dim nextInt As Integer = r.[Next](cpossibleChars.Length)
                Dim c As Char = cpossibleChars(nextInt)
                builder.Append(c)
            Next

            If CheckBox1.Checked = True Then
                Me.TextBox2.Text = builder.ToString().ToUpper
            ElseIf CheckBox3.Checked = True Then
                builder.Insert(5, "-")
                builder.Insert(11, "-")
                'builder.Insert(17, "-")
                'builder.Insert(23, "-")
                'builder.Insert(29, "-")
                Me.TextBox2.Text = builder.ToString.ToUpper
            End If

            TextBox1.Text += builder.ToString + Environment.NewLine

        Catch ex As Exception
            MessageBox.Show(String.Format("An error has occurred while trying to generate random password! Technical description: {0}", ex.Message.ToString()))
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Clipboard.SetText(TextBox1.Text)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        ProgressBar1.Value = 0
        TextBox3.Text = ""
        CheckBox1.Checked = False
        CheckBox3.Checked = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Timer1.Start()
        If TextBox3.Text = "" Then TextBox3.Text = "10"
        ProgressBar1.Maximum = TextBox3.Text
        If ProgressBar1.Value = ProgressBar1.Maximum Then ProgressBar1.Value = 0
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try
            If TextBox3.Text = "" Then TextBox3.Text = "10"

            If TextBox1.Lines.ToList.Count = TextBox3.Text Then
                Timer1.Stop()
            End If

            Dim cpossibleChars() As Char
            cpossibleChars = possibleChars.ToCharArray()
            If cpossibleChars.Length < 1 Then
                MessageBox.Show("You must enter one or more possible characters.")
                Return
            End If
            If len < 4 Then
                MessageBox.Show(String.Format("Please choose a password length. That length must be a value between {0} and {1}. Note: values above 1,000 might take a LONG TIME to process on some computers.", 4, Int32.MaxValue))
                Return
            End If

            Dim builder As New StringBuilder()

            For i As Integer = 0 To len - 1
                Dim randInt32 As Integer = GetRandomInt32Value.GetRandomInt()
                Dim r As New Random(randInt32)

                Dim nextInt As Integer = r.[Next](cpossibleChars.Length)
                Dim c As Char = cpossibleChars(nextInt)
                builder.Append(c)
            Next

            If CheckBox1.Checked = True Then
                Me.TextBox2.Text = builder.ToString().ToUpper()
            ElseIf CheckBox3.Checked = True Then
                builder.Insert(5, "-")
                builder.Insert(11, "-")
                'builder.Insert(17, "-")
                'builder.Insert(23, "-")
                'builder.Insert(29, "-")
                Me.TextBox2.Text = builder.ToString.ToUpper
            End If

            TextBox1.Text += builder.ToString.ToUpper + Environment.NewLine

            ProgressBar1.Value += 1

        Catch ex As Exception
            MessageBox.Show(String.Format("An error has occurred while trying to generate tokens! Technical description: {0}", ex.Message.ToString()))
            If ex.ToString.Contains("101") Then
                Timer1.Stop()
                ProgressBar1.Value = 0
                MessageBox.Show("An error has occurred and the operation was cancled!")
            End If
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Timer1.Stop()
        ProgressBar1.Value = 0
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox3.Checked = True Then CheckBox3.Checked = False
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox1.Checked = True Then CheckBox1.Checked = False
    End Sub
End Class

Public Class PasswordGenProfiler
    Public Shared Function GetFrequencyDistributionOfChars(allowableChars As String, generatedPass As String) As Dictionary(Of Char, Integer)
        Dim distrib As New Dictionary(Of Char, Integer)()
        ' initialize all values to 0
        For Each c As Char In allowableChars
            ' If character is listed more than once, don't re-add it to our list.
            If Not distrib.ContainsKey(c) Then
                distrib.Add(c, 0)
            End If
        Next
        Dim val As Integer = 0
        For Each passChar As Char In generatedPass
            If distrib.TryGetValue(passChar, val) Then
                distrib(passChar) = System.Threading.Interlocked.Increment(val)
            End If
        Next

        Return distrib
    End Function
End Class

Public Class RandomInt32Value
    Public Function GetRandomInt() As Integer
        Dim randomBytes As Byte() = New Byte(3) {}
        Dim rng As New RNGCryptoServiceProvider()
        rng.GetBytes(randomBytes)
        Dim randomInt As Integer = BitConverter.ToInt32(randomBytes, 0)
        Return randomInt
    End Function
End Class
