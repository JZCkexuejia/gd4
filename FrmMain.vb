
Imports System.Data.SqlClient

Public Class FrmMain
    'Dim sqlHelper As New SqlHelper
    ' Dim connStr As String = "Server=192.9.200.196;Database=DuvetDB;User Id=sa;Password=Aa123;"
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnRefresh_Click(vbNull, Nothing)
    End Sub

    Private Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        Dim connStr = txtConnStr.Text
        Dim tableName = txtTableName.Text
        Dim name = txtName.Text
        If name.Length > 0 Then
            Dim dt As DataTable = dgv.DataSource
            Dim drs As DataRow() = dt.Select("name = '" + name + "'")
            Dim sql As String
            If drs IsNot Nothing And drs.Length > 0 Then
                If MessageBox.Show("Row is exist,Are you sure you want to replace it", "", MessageBoxButtons.YesNo) = DialogResult.No Then
                    Return
                End If
                sql = "update " + tableName + " set "
                For index = 0 To dt.Columns.Count - 1
                    If ("ID".Equals(dt.Columns(index).ColumnName, StringComparison.CurrentCultureIgnoreCase) Or "Name".Equals(dt.Columns(index).ColumnName, StringComparison.CurrentCultureIgnoreCase)) Then
                        Continue For
                    End If
                    sql &= dt.Columns(index).ColumnName + " = '' , "
                Next
                sql = sql.Substring(0, sql.LastIndexOf(","))
                sql &= " where name = @name"
            Else
                sql = "insert into " + tableName + "(name) values(@name)"
            End If
            Dim pms = New SqlParameter() {New SqlParameter("@name", name)}
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand(sql, conn)
                    With cmd
                        .CommandType = CommandType.Text
                        .Parameters.AddRange(pms)
                        .ExecuteNonQuery()
                    End With
                    btnRefresh_Click(vbNull, Nothing)
                End Using
            End Using
        End If
    End Sub

    Private Sub btnAddColumns_Click(sender As Object, e As EventArgs) Handles btnAddColumns.Click
        Dim tableName = txtTableName.Text
        Dim connStr = txtConnStr.Text
        Dim colMin = txtColMin.Text + "_Min"
        Dim colMax = txtColMax.Text + "_Max"
        If colMin.Length > 0 And colMax.Length > 0 Then
            Dim dt As DataTable = dgv.DataSource
            If dt.Columns.Contains(colMin) Or dt.Columns.Contains(colMax) Then
                If MessageBox.Show("Columns is exist,Are you sure you want to replace it", "", MessageBoxButtons.YesNo) = DialogResult.No Then
                    Return
                End If
            End If
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand()
                    Dim sql As String
                    If dt.Columns.Contains(colMin) Then
                        sql = "update " + tableName + " set " + colMin + " = ''"
                    Else
                        sql = "alter table " + tableName + " add " + colMin + " varchar(64)"
                    End If
                    With cmd
                        .Connection = conn
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .ExecuteNonQuery()
                    End With
                    If dt.Columns.Contains(colMax) Then
                        sql = "update " + tableName + " set " + colMax + " = ''"
                    Else
                        sql = "alter table " + tableName + " add " + colMax + " varchar(64)"
                    End If
                    With cmd
                        .Connection = conn
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .ExecuteNonQuery()
                    End With
                End Using
            End Using
            btnRefresh_Click(vbNull, Nothing)
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Dim connStr = txtConnStr.Text
        Dim tableName = txtTableName.Text
        Dim sql = "select * from " + tableName
        Dim dt As New DataTable
        Using adapter As New SqlDataAdapter(sql, connStr)
            adapter.SelectCommand.CommandType = CommandType.Text
            adapter.Fill(dt)
        End Using
        If dt.Rows.Count > 0 Then
            dgv.DataSource = dt
        End If
    End Sub

    Private Sub btnModifyConnStr_Click(sender As Object, e As EventArgs) Handles btnModifyConnStr.Click
        'sqlHelper.connStr = txtConnStr.Text
        'sqlHelper.TableName = txtTableName.Text
    End Sub
End Class
