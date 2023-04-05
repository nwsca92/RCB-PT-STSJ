
Imports System.Data.OleDb
Imports System.IO
Imports System.Windows
Imports System.Windows.Forms.VisualStyles
Imports Microsoft.Office.Interop.Excel
Imports OleDbConnection = System.Data.OleDb.OleDbConnection

Public Class MasterLocation
    Dim conn As OleDbConnection
    Dim dta As OleDbDataAdapter
    Dim dts As DataSet
    Dim excel As String
    Dim OpenFileDialog As New OpenFileDialog


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            OpenFileDialog.Filter = "All Files (*.*)|*.*|Exccel Files (*.xlsx)|*.xls|Xls Files (*.xls)|*.xls"
            If OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                Dim fi As New IO.FileInfo(OpenFileDialog.FileName)
                Dim FileName As String = OpenFileDialog.FileName
                excel = fi.FullName
                conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excel + ";Extended Properties=Excel 12.0;")
                dta = New OleDbDataAdapter("select * from [Sheet1$]", conn)
                dts = New DataSet
                dta.Fill(dts, "[Sheet1$]")
                DataGridView1.DataSource = dts
                DataGridView1.DataMember = "[Sheet1$]"
                Label3.Text = DataGridView1.RowCount & " Records"
                TextBox1.Text = FileName
                conn.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
            Exit Sub
        End Try
    End Sub

End Class