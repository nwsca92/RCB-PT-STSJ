Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Forms.DataFormats

Public Class MainMenu
    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        Me.Close()
        Form1.Show()

        Form1.TextBox1.Text = ""
        Form1.TextBox2.Text = ""
        'Perintah ini digunakan untuk mengosongkan text1 dan text2 yang ada di form1.

        Form1.TextBox1.Focus()
        'Merupakan perintah untuk mengembalikan posisi kursor ke Text1 yang ada di form1.
    End Sub

    Private Sub MasterLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterLocationToolStripMenuItem.Click
        Dim newMDIChild As New MasterLocation()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub

    Private Sub MasterPartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterPartToolStripMenuItem.Click
        Dim newMDIChild As New MasterPart()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub RandomCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomCheckToolStripMenuItem.Click
        Dim newMDIChild As New RandomCheck()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel1.Text = Date.Now.ToString("dd MMMM yyyy HH:mm:ss")
    End Sub

    Private Sub StockOpnameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StockOpnameToolStripMenuItem.Click
        Dim newMDIChild As New StockOpname()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub CreditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreditToolStripMenuItem.Click
        Dim newMDIChild As New Credit()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub


End Class