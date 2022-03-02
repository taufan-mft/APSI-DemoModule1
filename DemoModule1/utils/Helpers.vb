Imports System.Data.OleDb

Module Helpers
    Sub clearForm(ParamArray var() As TextBox)
        For i As Integer = 0 To UBound(var, 1)
            var(i).Clear()
        Next
    End Sub

    Sub disableForm(ParamArray var() As TextBox)
        For i As Integer = 0 To UBound(var, 1)
            var(i).Enabled = False
        Next i
    End Sub

    Sub enableForm(ParamArray var() As TextBox)
        For i As Integer = 0 To UBound(var, 1)
            var(i).Enabled = True
        Next i
    End Sub

    Sub showDataToGrid(sequel As String, DGV As DataGridView)
        DA = New OleDb.OleDbDataAdapter(sequel, Conn)
        DS = New DataSet
        DA.Fill(DS)

        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Sub saveData(tableName As String, ParamArray var() As TextBox)
        Dim sql As String = "insert into " + tableName + " values("
        For i As Integer = 0 To UBound(var, 1)
            If i <> UBound(var, 1) Then
                sql = sql + "'" + var(i).Text + "',"
            Else
                sql = sql + "'" + var(i).Text + "')"
            End If

        Next
        CMD = New OleDb.OleDbCommand(sql, Conn)
        CMD.ExecuteNonQuery()
        clearForm(var)

    End Sub

    Sub deleteData(namatabel As String, namaid As String, id As String)
        Dim sql As String
        sql = "DELETE FROM " + namatabel + " WHERE " + namaid + " =" + id + ""
        CMD = New OleDb.OleDbCommand(sql, Conn)
        DM = CMD.ExecuteReader
        MsgBox("Data terhapus.")
    End Sub

    Sub updateData(tableName As String, idName As String, id As String, ParamArray var() As String)
        Dim sql As String
        sql = "update " + tableName + " set "
        For i As Integer = 0 To UBound(var, 1) Step 2
            If i <> (UBound(var, 1) - 1) Then
                sql = sql + var(i) + " ='" + var(i + 1) + "', "

            Else
                sql = sql + var(i) + " ='" + var(i + 1) + "'"
            End If
        Next
        sql = sql + " where " + idName + " = " + id + ""

        CMD = New OleDbCommand(sql, Conn)
        DM = CMD.ExecuteReader


    End Sub

    Function checkEmpty(ParamArray var() As TextBox) As Boolean
        Dim nomor As Integer = 0
        For i As Integer = 0 To UBound(var, 1)
            If (var(i).Text = "") Then
                nomor += 1
            End If
        Next
        Return nomor > 0
    End Function

    Function checkDuplicate(tableName As String, idColumnName As String, id As String)
        Dim sequel As String
        sequel = "select * from " + tableName + " where " + idColumnName + " = " + id + ""
        CMD = New OleDb.OleDbCommand(sequel, Conn)

        DM = CMD.ExecuteReader()
        DM.Read()

        If Not DM.HasRows Then
            Return False
        Else
            Return True
        End If

    End Function

    Sub showtoBox(row As Integer, DGV As DataGridView, ParamArray var() As TextBox)
        On Error Resume Next
        For i As Integer = 0 To UBound(var, 1)


            var(i).Text = DGV.Rows(row).Cells(i).Value

        Next
    End Sub

End Module
