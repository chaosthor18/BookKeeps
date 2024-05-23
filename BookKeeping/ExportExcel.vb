Imports Excel = Microsoft.Office.Interop.Excel
Public Class ExportExcel
    Private col_headerCount As Integer
    Private row_count As Integer
    Private col_headerT As New List(Of String)
    Private row_items As New List(Of String)
    Public Sub New(ByVal rowc As Integer, ByVal colh As Integer)
        col_headerCount = colh
        row_count = rowc
    End Sub
    Protected Overrides Sub Finalize()  ' destructor
    End Sub
    Public Sub add_headerTitle(ByVal val As String)
        col_headerT.Add(val)
    End Sub
    Public Sub col_headerTitle(ByVal ParamArray args() As String)
        For i As Integer = 0 To UBound(args, 1)
            col_headerT.Add(args(i))
        Next
    End Sub
    Public Sub addrow_Items(ByVal items As String)
        row_items.Add(items)
    End Sub
    Public Sub export(ByVal fileAddress As String)
        Dim xlapp As Excel.Application
        Dim xlworkbook As Excel.Workbook
        Dim xlworksheet As Excel.Worksheet
        Dim misvalue As Object = System.Reflection.Missing.Value
        Dim count_item As Integer = 0
        xlapp = New Excel.Application
        xlworkbook = xlapp.Workbooks.Add(misvalue)
        xlworksheet = xlworkbook.Sheets("sheet1")
        For col As Integer = 0 To col_headerCount - 1
            xlworksheet.Cells(1, col + 1) = col_headerT.Item(col)
            xlworksheet.Cells(1, col + 1) = col_headerT.Item(col)
        Next
        For row As Integer = 0 To row_count - 1
            For col As Integer = 0 To col_headerCount - 1
                xlworksheet.Cells(row + 2, col + 1) = row_items.Item(count_item)
                count_item += 1
            Next
        Next
        xlworksheet.SaveAs(fileAddress)
        xlworkbook.Close()
        xlapp.Quit()
        releaseObject(xlapp)
        releaseObject(xlworkbook)
        releaseObject(xlworksheet)
    End Sub
    Public Sub export_customDT(ByVal fileAddress As String)
        Dim xlapp As Excel.Application
        Dim xlworkbook As Excel.Workbook
        Dim xlworksheet As Excel.Worksheet
        Dim formatRange As Excel.Range
        Dim misvalue As Object = System.Reflection.Missing.Value
        Dim count_item As Integer = 0
        xlapp = New Excel.Application
        xlworkbook = xlapp.Workbooks.Add(misvalue)
        xlworksheet = xlworkbook.Sheets("sheet1")
        formatRange = xlworksheet.Range("B2", "B99000")
        formatRange.NumberFormat = "MMM-yyyy"
        For col As Integer = 0 To col_headerCount - 1
            xlworksheet.Cells(1, col + 1) = col_headerT.Item(col)
            xlworksheet.Cells(1, col + 1) = col_headerT.Item(col)
        Next
        For row As Integer = 0 To row_count - 1
            For col As Integer = 0 To col_headerCount - 1
                xlworksheet.Cells(row + 2, col + 1) = row_items.Item(count_item)
                count_item += 1
            Next
        Next
        xlworksheet.SaveAs(fileAddress)
        xlworkbook.Close()
        xlapp.Quit()
        releaseObject(xlapp)
        releaseObject(xlworkbook)
        releaseObject(xlworksheet)
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
End Class
