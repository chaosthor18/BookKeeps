Public Class ImportExcel
    Function import(ByVal fileAddress As String)
        Dim connection As System.Data.OleDb.OleDbConnection
        Dim dataset As System.Data.DataSet
        Dim command As System.Data.OleDb.OleDbDataAdapter

        connection = New System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" + fileAddress + "; Extended Properties=Excel 12.0;")
        command = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", connection)

        dataset = New System.Data.DataSet
        command.Fill(dataset)
        connection.Close()
        Return dataset
    End Function
End Class
