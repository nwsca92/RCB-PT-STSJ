Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1
    Public Shared savedusername As String = Nothing
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "admin" And TextBox2.Text = "admin" Then
            'MsgBox("Log in Successfully!", MsgBoxStyle.OkOnly, "Log in Form")
            MainMenu.Show()
            Me.Hide()
        ElseIf TextBox1.Text = "user" And TextBox2.Text = "user" Then
            MainMenu.Show()
            Me.Hide()
            'MainMenu.MenuStrip1.dropdown(1).Visible = False

        Else
            MsgBox("Incorrect Username and Password", MsgBoxStyle.OkOnly, "Warning")
        End If

        Form1.savedusername = TextBox1.Text
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Close()
    End Sub
End Class
