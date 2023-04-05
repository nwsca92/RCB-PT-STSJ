Imports System.Data.OleDb
Imports OleDbConnection = System.Data.OleDb.OleDbConnection

Public Class RandomCheck
    Dim pool As String = "0123456789"

    Dim conn As OleDbConnection
    Dim dta As OleDbDataAdapter
    Dim dts As DataSet
    Dim excel As String
    Dim OpenFileDialog As New OpenFileDialog

    Private Sub RandomCheck_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True

        Dim count = 0
        TextBox7.Text = ""

        Dim cc As New Random
        Dim strpos = ""
        While count <= 16
            strpos = cc.Next(0, pool.Length)
            TextBox7.Text = TextBox7.Text & pool(strpos)
            count = count + 1
        End While
        TextBox6.Text = Form1.savedusername
        TextBox3.Text = Date.Now.ToString("MMMM yyyy")
        'TextBox4.Text = Date.Now.ToString("dd MMMM yyyy HH:MM:ss")
        'Label14.Text = Date.Now.ToString("dd MMMM yyyy HH:MM:ss")
        'DataGridView2.DataSource = MasterLocation.DataGridView1.DataSource
        DataGridView1.Columns(0).HeaderCell.Style.BackColor = Color.Yellow
        DataGridView1.Columns(1).HeaderCell.Style.BackColor = Color.Yellow
        DataGridView1.Columns(2).HeaderCell.Style.BackColor = Color.Yellow
        DataGridView1.Columns(3).HeaderCell.Style.BackColor = Color.Yellow
        DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.LightBlue
        DataGridView1.Columns(5).HeaderCell.Style.BackColor = Color.LightBlue
        DataGridView1.Columns(6).HeaderCell.Style.BackColor = Color.LightBlue
        'DataGridView1.Columns(7).HeaderCell.Style.BackColor = Color.LightBlue
        ComboBox2.SelectedIndex = "0"

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim num1 As Double
        Dim num2 As Double
        Dim num3 As Double

        If TextBox5.Text = "" Then
            MessageBox.Show("Fill Qty")
            TextBox5.Focus()

        Else
            num1 = Convert.ToDouble(TextBox5.Text)
            num2 = Convert.ToDouble(TextBox8.Text)
            num3 = num1 - num2

            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(2).Value = TextBox1.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(0).Value = TextBox2.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(3).Value = TextBox5.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1).Value = Label11.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(4).Value = Label12.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(5).Value = TextBox8.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(6).Value = CStr(num3)
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(7).Value = Label14.Text
            DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(8).Value = TextBox4.Text
            DataGridView1.Update()

            'TextBox2.Text = ""
            'TextBox2.Focus()
            TextBox5.Text = ""
            Label15.Text = DataGridView1.RowCount - 1
        End If
    End Sub



    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        Dim allowedChars As String = "0123456789"
        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

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
                'DataGridView2.DataSource = dts
                DataGridView2.DataMember = "[Sheet1$]"

                'Dim bs As New BindingSource
                BindingSource1.DataSource = dts
                'bs.DataSource = dts
                'BindingSource1.DataSource = dts
                DataGridView2.DataSource = BindingSource1

                For Each row In DataGridView2.Rows
                    ComboBox1.Items.Add((row.Cells(3)).Value).ToString()
                Next

                conn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
            Exit Sub
        End Try
    End Sub

    Private Sub DataGridView2_GotFocus(sender As Object, e As EventArgs) Handles DataGridView2.GotFocus

        If DataGridView2.SelectedRows.Count = 0 Then
            'MessageBox.Show("not found")
            TextBox2.Text = ""
            TextBox2.Focus()
        Else Dim current_row As Integer = DataGridView2.CurrentRow.Index
            Debug.Print(current_row.ToString)
            Label11.Text = DataGridView2(2, current_row).Value.ToString
            Label12.Text = DataGridView2(3, current_row).Value.ToString
            Label13.Text = DataGridView2(4, current_row).Value.ToString
            TextBox8.Text = DataGridView2(4, current_row).Value.ToString
        End If
    End Sub

    Private Sub TextBox2_Enter(sender As Object, e As EventArgs) Handles TextBox2.Enter
        Dim dsView As New DataView()
        Dim dts As New DataSet
        'Dim bs As New BindingSource
        Try
            conn.Open()
            dts.Clear()
            dta.Fill(dts, "[Sheet1$]")
            dsView = dts.Tables(0).DefaultView
            BindingSource1.DataSource = dsView
            DataGridView2.DataSource = BindingSource1
            BindingSource1.Filter = "[Parts_Code] LIKE '%" & TextBox2.Text & "%'"

        Catch ex As Exception
            MessageBox.Show("(1) Nothing Found In My List, Try Again.. ")
            BindingSource1.RemoveFilter()
        Finally
            DataGridView2.DataSource = BindingSource1
            conn.Close()
        End Try
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyValue = Keys.Enter Then
            Dim dsView As New DataView()
            Dim dts As New DataSet

            Try
                conn.Open()
                dts.Clear()
                dta.Fill(dts, "[Sheet1$]")
                dsView = dts.Tables(0).DefaultView
                BindingSource1.DataSource = dsView
                DataGridView2.DataSource = BindingSource1
                BindingSource1.Filter = "[Parts_Code] LIKE '%" & TextBox2.Text & "%'"
                'DataGridView2.Focus()
                DataGridView2.Select()

            Catch ex As Exception
                MessageBox.Show("(2) Nothing Found In My List, Try Again.. ")
                BindingSource1.RemoveFilter()
            Finally
                DataGridView2.DataSource = BindingSource1
                conn.Close()
            End Try
        End If
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.Name = TextBox7.Text


        'For HEADERS
        xlWorkSheet.Cells(1, 3) = "Random Check Barcode"
        xlWorkSheet.Cells(2, 3) = ComboBox2.SelectedText


        For i = 1 To DataGridView1.ColumnCount
            xlWorkSheet.Cells(1, i) = Me.DataGridView1.Columns(i - 1).HeaderText
            'FOR ITEMS
            For j = 1 To DataGridView1.RowCount
                xlWorkSheet.Cells(j + 1, i) = Me.DataGridView1(i - 1, j - 1).Value
            Next
        Next


        Dim strFileDestination As String
        With SaveFileDialog1
            .Filter = "Excel Office (.xls)|*.xlsx"
            .FileName = "RCB ID = " + TextBox7.Text
            .ShowDialog()
            strFileDestination = .FileName
        End With

        xlWorkBook.SaveAs(strFileDestination)
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'TextBox4.Text = Date.Now.ToString("dd MMMM yyyy HH:mm:ss")
        TextBox4.Text = Format(Now, "dd MMMM yyyy HH:mm")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        For Each r As DataGridViewRow In DataGridView2.Rows
            'If r.IsNewRow Then Continue For

            'r.Cells(0).Value is the current row's first column, r.Cells(1).Value is the second column
            DataGridView1.Rows.Add({"", "", "", "", r.Cells(3).Value, r.Cells(4).Value})
            'ComboBox1.Items.Add({r.Cells(3).Value})
        Next
    End Sub
    Private Sub DataGridView2_Paint(sender As Object, e As PaintEventArgs) _
    Handles DataGridView2.Paint
        If DataGridView2.Rows.Count = 0 Then
            TextBox2.Enabled = False
            'TextRenderer.DrawText(e.Graphics, "No records found.",
            'DataGridView2.Font, DataGridView1.ClientRectangle,
            'DataGridView2.ForeColor, DataGridView2.BackgroundColor,
            'TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        Else
            TextBox2.Enabled = True
            'TextBox2.Focus()
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Label11.Text = "-"
        Label12.Text = "-"
        Label13.Text = "-"
        Label14.Text = "-"
        TextBox2.Text = ""
        TextBox8.Text = "0"
        TextBox2.Focus()
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyValue = Keys.Enter Then
            Dim num1 As Double
            Dim num2 As Double
            Dim num3 As Double

            If TextBox5.Text = "" Then
                MessageBox.Show("Fill Qty")
                TextBox5.Focus()
            Else
                num1 = Convert.ToDouble(TextBox5.Text)
                num2 = Convert.ToDouble(TextBox8.Text)
                num3 = num1 - num2

                DataGridView1.Rows.Add(1)
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(2).Value = TextBox1.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(0).Value = TextBox2.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(3).Value = TextBox5.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1).Value = Label11.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(4).Value = Label12.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(5).Value = TextBox8.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(6).Value = CStr(num3)
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(7).Value = Label14.Text
                DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(8).Value = TextBox4.Text
                DataGridView1.Update()

                'TextBox2.Text = ""
                'TextBox2.Focus()
                TextBox5.Text = ""
                Label15.Text = DataGridView1.RowCount - 1
            End If

        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'Dim dsView As New DataView()
        'Dim dts As New DataSet

        'Try
        '    conn.Open()
        '    dts.Clear()
        '    dta.Fill(dts, "[Sheet1$]")
        '    dsView = dts.Tables(0).DefaultView
        '    BindingSource1.DataSource = dsView
        '    DataGridView2.DataSource = BindingSource1
        '    BindingSource1.Filter = "[Parts_Code] LIKE '%" & TextBox2.Text & "%'"
        '    'DataGridView2.Focus()
        '    'DataGridView2.Select()
        'Catch ex As Exception
        '    MessageBox.Show("Nothing Found")
        'Finally
        '    conn.Close()
        'End Try
    End Sub

    Private Sub DataGridView1_Paint(sender As Object, e As PaintEventArgs) Handles DataGridView1.Paint
        If DataGridView1.Rows.Count = 1 Then
            Button5.Enabled = False

            'TextRenderer.DrawText(e.Graphics, "No records found.",
            'DataGridView2.Font, DataGridView1.ClientRectangle,
            'DataGridView2.ForeColor, DataGridView2.BackgroundColor,
            'TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        Else
            Button5.Enabled = True
            'TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView2.Refresh()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Dim refA As String = Label16.Text
        'Dim refB As String = Label13.Text
        'If refA.Equals(refB) Then
        '    Label14.Text = "OK"
        'Else
        '    Label14.Text = "MOVING"
        'End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.SelectedItem
    End Sub
End Class

