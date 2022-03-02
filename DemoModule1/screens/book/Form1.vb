Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksiDB()
        showDataToGrid($"SELECT * FROM {TABLE_BOOK}", DataGridView1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not checkEmpty(TextBox1, TextBox2, TextBox3) Then
            If Not checkDuplicate(TABLE_BOOK, "id", TextBox1.Text) Then
                saveData(TABLE_BOOK, TextBox1, TextBox2, TextBox3)
            Else
                MsgBox("Duplicates Found")
            End If
        End If
        showDataToGrid($"SELECT * FROM {TABLE_BOOK}", DataGridView1)
    End Sub

    Private Sub DGV_MouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        showtoBox(e.RowIndex, DataGridView1, TextBox1, TextBox2, TextBox3)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        deleteData(TABLE_BOOK, "id", TextBox1.Text)
        clearForm(TextBox1, TextBox2, TextBox3)
        showDataToGrid($"SELECT * FROM {TABLE_BOOK}", DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        updateData(TABLE_BOOK, "id", TextBox1.Text, "book_title", TextBox2.Text, "issbn", TextBox3.Text)
        clearForm(TextBox1, TextBox2, TextBox3)
        showDataToGrid($"SELECT * FROM {TABLE_BOOK}", DataGridView1)
    End Sub
End Class
