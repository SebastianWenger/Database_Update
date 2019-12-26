Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class SQLDatabase
    Public Sub Insert(dt As DataTable)
        Dim connstring As String = "Data Source = tcp:cltsaporder-c30,49172; Initial Catalog = Master; Integrated Security = True;"
        Using cn As SqlConnection = New SqlConnection(connstring)
            cn.Open()
            Using bulkCopy As SqlClient.SqlBulkCopy = New SqlClient.SqlBulkCopy(cn)
                bulkCopy.DestinationTableName = "dbo.ZSD_CONT_LIST"
                bulkCopy.WriteToServer(dt)

            End Using
        End Using
    End Sub

    Public Sub RemoveDuplicates()
        Dim sqlquery As String = "WITH ToDelete AS (SELECT ROW_NUMBER() OVER (PARTITION BY [Sales Document],[Contract Line Item] ORDER BY [Updated On] DESC) AS rn FROM dbo.ZSD_CONT_LIST)
                                  DELETE FROM ToDelete
                                  WHERE rn > 1"

        Dim connstring As String = "Data Source = tcp:cltsaporder-c30,49172; Initial Catalog = Master; Integrated Security = True;"
        'Dim connstring As String = "Data Source = SWENGER-5480\SQLEXPRESS; Initial Catalog = Master; Integrated Security = True;"
        Using cn As SqlConnection = New SqlConnection(connstring)
            cn.Open()
            Using cmd As SqlCommand = New SqlCommand(sqlquery, cn)
                cmd.ExecuteNonQuery()
            End Using
            cn.Close()
        End Using
    End Sub

    Public Function ImportExceltoDatatable(ByVal filepath As String) As DataTable

        Dim sqlquery As String = "Select * From [Sheet1$]"
        Dim ds As DataSet = New DataSet()
        Dim constring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";Extended Properties=""Excel 12.0;HDR=YES;"""
        Dim con As OleDbConnection = New OleDbConnection(constring & "")
        Dim da As OleDbDataAdapter = New OleDbDataAdapter(sqlquery, con)
        da.Fill(ds)
        Dim dt As DataTable = ds.Tables(0)
        Return dt
    End Function

    Public Sub AddDateColumn(dt As DataTable)
        Dim colDateTime As New DataColumn("Updated On", GetType(DateTime))
        colDateTime.DefaultValue = Now().ToString()
        dt.Columns.Add(colDateTime)
    End Sub


End Class
