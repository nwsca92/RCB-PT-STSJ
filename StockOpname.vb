Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar

Public Class StockOpname
    Private Sub StockOpname_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim pool As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim cc As New Random
        Dim strpos = ""
        Dim count = 0
        TextBox1.Text = ""
        While count <= 16
            strpos = cc.Next(0, pool.Length)
            TextBox1.Text = TextBox1.Text & pool(strpos)
            count = count + 1
        End While

        TextBox3.Text = Form1.savedusername
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = MessageBox.Show("Finish Current SO?", "Stock Opname", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            MessageBox.Show("SO Finished")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Go Back")
        End If
    End Sub
End Class