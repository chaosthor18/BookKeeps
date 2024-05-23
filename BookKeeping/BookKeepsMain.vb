Imports System.IO

Public Class frmMain
    Private Property MoveForm As Boolean
    Private Property MoveForm_MousePosition As Point

    'for cash receipt book
    Private number As Integer = 0
    Private or_number As Integer = 0
    Private merchandiseT As Double = 0
    Private chargesT As Double = 0
    Private actualCashT As Double = 0

    'for cash disbursement book
    Private creditTotal_Cdr As Double
    Private debitTotal_Cdr As Double
    Private selected_columnCdr As String
    Private items_comboboxCdb As New Stack(Of String)
    Private items_valuescbCdb As New Dictionary(Of String, Double)

    Private Sub frmMain_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown, pnlControl.MouseDown
        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.Default
            MoveForm_MousePosition = e.Location
        End If
    End Sub
    Private Sub frmMain_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp, pnlControl.MouseUp
        If e.Button = MouseButtons.Left Then
            MoveForm = False
            Me.Cursor = Cursors.Default
        End If
    End Sub
    Private Sub frmMain_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove, pnlControl.MouseMove
        If MoveForm Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AssignValidation(Me.txtboxMerchandise_CRB, ValidationType.Only_Numbers)
        AssignValidation(Me.txtboxCharge_CRB, ValidationType.Only_Numbers)
        AssignValidation(Me.txtboxNum_CRB, ValidationType.Integer_Number)
        AssignValidation(Me.txtboxORnum_CRB, ValidationType.Integer_Number)
        AssignValidation(Me.txtboxCash_Cdb, ValidationType.Only_Numbers)
        AssignValidation(Me.txtboxValue_Gj, ValidationType.Only_Numbers)
        AssignValidation(Me.txtboxDebit_L, ValidationType.Only_Numbers)
        AssignValidation(Me.txtboxCredit_L, ValidationType.Only_Numbers)
        dgvCashReceipt.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        dgvCashDisburse.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
        dgvJournal.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        dgvLedger.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        dtpDate_CRB.Format = DateTimePickerFormat.Custom
        dtpDate_CDB.Format = DateTimePickerFormat.Custom
        dtpDate_Gj.Format = DateTimePickerFormat.Custom
        dtpDate_CRB.CustomFormat = "MM/dd/yyyy"
        dtpDate_CDB.CustomFormat = "MM/dd/yyyy"
        dtpDate_Gj.CustomFormat = "MM/dd/yyyy"
        dtpDate_CDB.Value = Today
        dtpDate_CRB.Value = Today
        dtpDate_Gj.Value = Today
        cbMonth_Ledger.SelectedIndex = Microsoft.VisualBasic.Month(Today.Date) - 1
        cbMonth_Ledger.SelectedItem = Today.Year
        lblMerchandiseT_CRB.Text = merchandiseT
        lblTotalcharge_CRB.Text = chargesT
        lblActualcr_CRB.Text = actualCashT
        lblCreditTotal_Cdb.Text = creditTotal_Cdr
        lbldebitTotal_Cdb.Text = debitTotal_Cdr
        items_comboboxCdb.Push("Add Column")
        items_comboboxCdb.Push("Debit Inventory")
        col_notsortable()
        For year As Integer = 2010 To 2040
            cbYear_Ledger.Items.Add(year)
        Next
        For Each item In items_comboboxCdb
            cbSelectColumn_Cbr.Items.Add(item)
        Next
        cbYear_Ledger.Text = Microsoft.VisualBasic.DateAndTime.Year(Today).ToString
        pnlCrb.Visible = True
        pnlCdb.Visible = False
        pnlJournal.Visible = False
        pnlLedger.Visible = False
        pnlAbout.Visible = False
    End Sub


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
    Private Sub col_notsortable()
        For index As Integer = 0 To dgvCashDisburse.ColumnCount - 1
            dgvCashDisburse.Columns.Item(index).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For index As Integer = 0 To dgvCashReceipt.ColumnCount - 1
            dgvCashReceipt.Columns.Item(index).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub
    'Cash Receipt
    Private Sub add_ChargesMerchandise(ByVal value As String, ByVal add_categ As String)
        Dim numerical_value As Double
        If Double.TryParse(value, numerical_value) Then
            If (add_categ = "MERCHANDISE") Then
                merchandiseT += numerical_value
                actualCashT += numerical_value
                txtboxMerchandise_CRB.Text = ""
            ElseIf (numerical_value < actualCashT And add_categ = "CHARGES") Then
                chargesT += numerical_value
                actualCashT -= numerical_value
                txtboxCharge_CRB.Text = ""
            Else
                Dim msgbox1 As Message_Box = New Message_Box(0, "Charge is greater than cash received")
            End If
        ElseIf String.IsNullOrEmpty(value) Then
            Dim msgbox2 As Message_Box = New Message_Box(0, "Please enter a value")
        Else
            Dim msgbox3 As Message_Box = New Message_Box(0, "Enter numerical value only")
            txtboxMerchandise_CRB.Text = ""
            txtboxCharge_CRB.Text = ""
        End If
        updateValue()
    End Sub
    Private Sub updateValue()
        lblMerchandiseT_CRB.Text = merchandiseT
        lblActualcr_CRB.Text = actualCashT
        lblTotalcharge_CRB.Text = chargesT
    End Sub
    Private Sub btnaddM_Click(sender As Object, e As EventArgs) Handles btnaddM_CRB.Click
        add_ChargesMerchandise(txtboxMerchandise_CRB.Text, "MERCHANDISE")
    End Sub
    Private Sub btnaddC_Click(sender As Object, e As EventArgs) Handles btnaddC_CRB.Click
        add_ChargesMerchandise(txtboxCharge_CRB.Text, "CHARGES")
    End Sub
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit_CRB.Click
        Dim submit_data() As String = {txtboxNum_CRB.Text, txtboxORnum_CRB.Text, txtboxClientN_CRB.Text, dtpDate_CRB.Text}
        Dim name_data() As String = {"Number", "OR Number", "Client Name", "Date"}
        Dim count As Integer
        Dim count_empty As Integer = 0
        Dim error_list As String = ""
        For count = 0 To 3
            If (String.IsNullOrEmpty(submit_data(count))) Then
                count_empty += 1
                error_list += name_data(count) + "|"
            End If
        Next
        If count_empty = 0 And dgvCashReceipt.RowCount <> 0 Then
            If Not Double.TryParse(txtboxNum_CRB.Text, 0) Or Not Double.TryParse(txtboxORnum_CRB.Text, 0) Then
                Dim msgbox1 As Message_Box = New Message_Box(0, "Please check in Number or OR/SI Number if they are a numerical value")
                Return
            End If
            If (dgvCashReceipt(3, dgvCashReceipt.RowCount - 1).Value.ToString <> "Total:") Then
                dgvCashReceipt.Rows.Add(submit_data(0), submit_data(3), submit_data(2), submit_data(1), merchandiseT, chargesT, actualCashT)
                dtpDate_CRB.Text = dtpDate_CRB.Value.AddDays(1)
                txtboxNum_CRB.Text += 1
                txtboxORnum_CRB.Text += 1
            Else
                Dim msgbox2 As Message_Box = New Message_Box(0, "Error: The data has total")
            End If
        ElseIf count_empty = 0 Then
            dgvCashReceipt.Rows.Add(submit_data(0), submit_data(3), submit_data(2), submit_data(1), merchandiseT, chargesT, actualCashT)
            dtpDate_CRB.Text = dtpDate_CRB.Value.AddDays(1)
            txtboxNum_CRB.Text += 1
            txtboxORnum_CRB.Text += 1
        ElseIf count_empty > 0 Then
            Dim msgbox3 As Message_Box = New Message_Box(0, error_list & " is/are empty")
        End If
        merchandiseT = 0
        chargesT = 0
        actualCashT = 0
        updateValue()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvCashReceipt.RowCount <> 0 Then
            Dim filename As String
            filename = Microsoft.VisualBasic.DateAndTime.MonthName(Today.Month).ToString & " " & Microsoft.VisualBasic.DateAndTime.Year(Today).ToString & " " & "Cash Receipt Book"
            Using sfd As SaveFileDialog = New SaveFileDialog With {.Filter = "Excel Workbook|*.xlsx", .Title = "Export", .FileName = filename}
                Try
                    If sfd.ShowDialog() = DialogResult.OK Then
                        Dim exportxl As ExportExcel = New ExportExcel(dgvCashReceipt.RowCount, dgvCashReceipt.ColumnCount)
                        exportxl.col_headerTitle("Number", "Date", "Client Name", "OR/SI", "Merchandise Total", "Charges", "Actual Cash Received")
                        For row As Integer = 0 To dgvCashReceipt.RowCount - 1
                            For col As Integer = 0 To dgvCashReceipt.ColumnCount - 1
                                exportxl.addrow_Items(dgvCashReceipt(col, row).Value.ToString)
                            Next
                        Next
                        exportxl.export(sfd.FileName.ToString)
                        Dim msgbox1 As Message_Box = New Message_Box(1, "Successfully Exported")
                    End If
                Catch ex As Exception
                    Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message.ToString)
                End Try
            End Using
        Else
            Dim msgbox2 As Message_Box = New Message_Box(0, "The table is empty")
        End If
    End Sub
    Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub btnMinimize_Click(sender As Object, e As EventArgs) Handles btnMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub btnImport_CRB_Click(sender As Object, e As EventArgs) Handles btnImport_CRB.Click
        Using ofd As OpenFileDialog = New OpenFileDialog With {.Filter = "Excel Workbook|*.xlsx", .Title = "Import"}
            Try
                If ofd.ShowDialog() = DialogResult.OK Then
                    If Not (ofd.FileName.ToString.Contains("Cash Receipt Book")) Then
                        Dim msgbox1 As Message_Box = New Message_Box(0, "Please import only Cash Receipt Book")
                        Return
                    End If
                    Dim importxl As ImportExcel = New ImportExcel()
                    Dim fileAddress = New FileInfo(ofd.FileName)
                    Dim dataset As System.Data.DataSet

                    dgvCashReceipt.Rows.Clear()
                    dataset = importxl.import(fileAddress.ToString)
                    For row As Integer = 0 To dataset.Tables(0).Rows.Count - 1
                        Dim date_tostring As String = ""
                        Dim orsi_tostring As String = "Total:"
                        If Not IsDBNull(dataset.Tables(0).Rows(row).Item(1)) Then
                            date_tostring = dataset.Tables(0).Rows(row).Item(1)
                            orsi_tostring = dataset.Tables(0).Rows(row).Item(3)
                        End If
                        dgvCashReceipt.Rows.Add(dataset.Tables(0).Rows(row).Item(0), date_tostring, dataset.Tables(0).Rows(row).Item(2), orsi_tostring, dataset.Tables(0).Rows(row).Item(4), dataset.Tables(0).Rows(row).Item(5), dataset.Tables(0).Rows(row).Item(6))
                    Next
                    Dim msgbox2 As Message_Box = New Message_Box(1, "Import Complete")
                    col_notsortable()
                End If
            Catch ex As Exception
                Dim msgbox3 As Message_Box = New Message_Box(0, ex.Message.ToString)
            End Try
        End Using
    End Sub
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Dim col_head() As String = {"Number", "Date", "Client Name", "OR/SI", "Merchandise Total", "Charges", "Actual Cash Received"}
        col_notsortable()
        txtboxCharge_CRB.Text = ""
        txtboxClientN_CRB.Text = ""
        txtboxNum_CRB.Text = ""
        txtboxORnum_CRB.Text = ""
        txtboxMerchandise_CRB.Text = ""
        dtpDate_CRB.Text = Today
        merchandiseT = 0
        chargesT = 0
        actualCashT = 0
        lblActualcr_CRB.Text = 0
        lblMerchandiseT_CRB.Text = 0
        lblTotalcharge_CRB.Text = 0
        Dim style As DataGridViewCellStyle = dgvCashReceipt.ColumnHeadersDefaultCellStyle()
        dgvCashReceipt.Rows.Clear()
        dgvCashReceipt.ColumnHeadersDefaultCellStyle = style
    End Sub
    Private Sub btnTotal_Click(sender As Object, e As EventArgs) Handles btnTotal_CRB.Click
        If dgvCashReceipt.RowCount <> 0 Then
            If (dgvCashReceipt(3, dgvCashReceipt.RowCount - 1).Value.ToString <> "Total:") Then
                Dim total_merchandise As Double
                Dim total_charges As Double
                Dim total_actualcashR As Double
                For row As Integer = 0 To dgvCashReceipt.RowCount - 1
                    total_merchandise += dgvCashReceipt(4, row).Value
                Next
                For row As Integer = 0 To dgvCashReceipt.RowCount - 1
                    total_charges += dgvCashReceipt(5, row).Value
                Next
                For row As Integer = 0 To dgvCashReceipt.RowCount - 1
                    total_actualcashR += dgvCashReceipt(6, row).Value
                Next
                dgvCashReceipt.Rows.Add("", "", "", "Total:", total_merchandise, total_charges, total_actualcashR)
            Else
                Dim msgbox1 As Message_Box = New Message_Box(0, "The table has total")
            End If
        Else
            Dim msgbox2 As Message_Box = New Message_Box(0, "Empty table")
        End If
    End Sub
    Private Sub btnDelete_Crb_Click(sender As Object, e As EventArgs) Handles btnDelete_Crb.Click
        If Not (dgvCashReceipt.Rows.Count = 0) Then
            dgvCashReceipt.Rows.RemoveAt(dgvCashReceipt.Rows.Count - 1)
        Else
            Dim msgbox1 As Message_Box = New Message_Box(0, "You cannot delete a row in an empty table.")
        End If
    End Sub
    'Cash Receipt
    'Side Panel
    Private Sub btnCD_Click(sender As Object, e As EventArgs) Handles btnCD.Click
        pnlCrb.Visible = False
        pnlCdb.Visible = True
        pnlJournal.Visible = False
        pnlLedger.Visible = False
        pnlAbout.Visible = False
        btnCR.ForeColor = Color.White
        btnCR.BackColor = Color.Black
        btnCD.ForeColor = Color.Black
        btnCD.BackColor = Color.DimGray
        btnLedger.ForeColor = Color.White
        btnLedger.BackColor = Color.Black
        btnJournal.ForeColor = Color.White
        btnJournal.BackColor = Color.Black
        btnAbout.ForeColor = Color.White
        btnAbout.BackColor = Color.Black
    End Sub
    Private Sub btnCR_Click(sender As Object, e As EventArgs) Handles btnCR.Click
        pnlCrb.Visible = True
        pnlCdb.Visible = False
        pnlJournal.Visible = False
        pnlLedger.Visible = False
        pnlAbout.Visible = False
        btnCR.ForeColor = Color.Black
        btnCR.BackColor = Color.DimGray
        btnCD.ForeColor = Color.White
        btnCD.BackColor = Color.Black
        btnLedger.ForeColor = Color.White
        btnLedger.BackColor = Color.Black
        btnJournal.ForeColor = Color.White
        btnJournal.BackColor = Color.Black
        btnAbout.ForeColor = Color.White
        btnAbout.BackColor = Color.Black
    End Sub
    Private Sub btnJournal_Click(sender As Object, e As EventArgs) Handles btnJournal.Click
        pnlCrb.Visible = False
        pnlCdb.Visible = False
        pnlJournal.Visible = True
        pnlLedger.Visible = False
        pnlAbout.Visible = False
        btnCR.ForeColor = Color.White
        btnCR.BackColor = Color.Black
        btnCD.ForeColor = Color.White
        btnCD.BackColor = Color.Black
        btnLedger.ForeColor = Color.White
        btnLedger.BackColor = Color.Black
        btnAbout.ForeColor = Color.White
        btnAbout.BackColor = Color.Black
        btnJournal.ForeColor = Color.Black
        btnJournal.BackColor = Color.DimGray
    End Sub
    Private Sub btnLedger_Click(sender As Object, e As EventArgs) Handles btnLedger.Click
        pnlCrb.Visible = False
        pnlCdb.Visible = False
        pnlJournal.Visible = False
        pnlLedger.Visible = True
        pnlAbout.Visible = False
        btnCR.ForeColor = Color.White
        btnCR.BackColor = Color.Black
        btnCD.ForeColor = Color.White
        btnCD.BackColor = Color.Black
        btnJournal.ForeColor = Color.White
        btnJournal.BackColor = Color.Black
        btnLedger.ForeColor = Color.Black
        btnLedger.BackColor = Color.DimGray
        btnAbout.ForeColor = Color.White
        btnAbout.BackColor = Color.Black
    End Sub
    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        pnlCrb.Visible = False
        pnlCdb.Visible = False
        pnlJournal.Visible = False
        pnlLedger.Visible = False
        pnlAbout.Visible = True
        btnCR.ForeColor = Color.White
        btnCR.BackColor = Color.Black
        btnCD.ForeColor = Color.White
        btnCD.BackColor = Color.Black
        btnJournal.ForeColor = Color.White
        btnJournal.BackColor = Color.Black
        btnLedger.ForeColor = Color.White
        btnLedger.BackColor = Color.Black
        btnAbout.ForeColor = Color.Black
        btnAbout.BackColor = Color.DimGray
    End Sub
    'Side Panel
    'cash disbursement
    Private Sub cbSelectColumn_Cbr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSelectColumn_Cbr.SelectedIndexChanged
        selected_columnCdr = cbSelectColumn_Cbr.Text
        txtboxValue_Cdb.Clear()
    End Sub
    Private Sub btnaddCash_Cdb_Click(sender As Object, e As EventArgs) Handles btnaddCash_Cdb.Click
        Dim temp As Double
        If (Double.TryParse(txtboxCash_Cdb.Text, temp)) Then
            If (txtboxCash_Cdb.Text < 0) Then
                Dim msgbox1 As Message_Box = New Message_Box(0, "Enter Positive Numbers Only")
                Return
            End If
            creditTotal_Cdr += txtboxCash_Cdb.Text
            lblCreditTotal_Cdb.Text = creditTotal_Cdr
            lbldebitTotal_Cdb.Text = debitTotal_Cdr
        Else
            Dim msgbox2 As Message_Box = New Message_Box(0, "Enter Numbers Only")
        End If
        txtboxCash_Cdb.Text = 0
    End Sub
    Private Sub btnaddValue_Cdb_Click(sender As Object, e As EventArgs) Handles btnaddValue_Cdb.Click
        If Not String.IsNullOrWhiteSpace(txtboxValue_Cdb.Text) Then
            If (selected_columnCdr = "Add Column") Then
                If Not items_comboboxCdb.Contains("Debit " & txtboxValue_Cdb.Text) Then
                    dgvCashDisburse.Columns.Add("Column" & dgvCashDisburse.ColumnCount, "Debit" & " " & txtboxValue_Cdb.Text)
                    items_comboboxCdb.Push("Debit " & txtboxValue_Cdb.Text)
                    cbSelectColumn_Cbr.Items.Clear()
                    For Each item In items_comboboxCdb
                        cbSelectColumn_Cbr.Items.Add(item)
                    Next
                Else
                    Dim msgbox1 As Message_Box = New Message_Box(0, "Column is Existing")
                End If
                txtboxValue_Cdb.Text = ""
            ElseIf (String.IsNullOrWhiteSpace(selected_columnCdr)) Then
                Dim msgbox2 As Message_Box = New Message_Box(0, "Please select column")
            Else
                Dim val As Double
                If (Double.TryParse(txtboxValue_Cdb.Text, val)) Then
                    If txtboxValue_Cdb.Text < 0 Then
                        Dim msgbox3 As Message_Box = New Message_Box(0, "Please enter positive number only")
                        Return
                    End If
                    Dim result As String = Nothing
                    If items_valuescbCdb.TryGetValue(cbSelectColumn_Cbr.Text, result) Then
                        items_valuescbCdb(cbSelectColumn_Cbr.Text) = val + result
                    Else
                        items_valuescbCdb.Add(cbSelectColumn_Cbr.Text, val)
                    End If
                    debitTotal_Cdr += val
                    txtboxValue_Cdb.Text = ""
                Else
                    Dim msgbox4 As Message_Box = New Message_Box(0, "Please enter number only")
                End If
            End If
        Else
            Dim msgbox5 As Message_Box = New Message_Box(0, "Please enter value")
        End If
        cbSelectColumn_Cbr.SelectedIndex = cbSelectColumn_Cbr.FindStringExact("")
        selected_columnCdr = ""
        lblCreditTotal_Cdb.Text = creditTotal_Cdr
        lbldebitTotal_Cdb.Text = debitTotal_Cdr
    End Sub
    Private Sub insertrow_Cdb()
        dgvCashDisburse.Rows.Add()
        dgvCashDisburse(0, dgvCashDisburse.RowCount - 1).Value = dtpDate_CDB.Text
        dgvCashDisburse(1, dgvCashDisburse.RowCount - 1).Value = txtboxSupplierN_Cdb.Text
        dgvCashDisburse(2, dgvCashDisburse.RowCount - 1).Value = txtboxRef_Cdb.Text
        dgvCashDisburse(3, dgvCashDisburse.RowCount - 1).Value = creditTotal_Cdr
        For col As Integer = 4 To dgvCashDisburse.ColumnCount - 1
            Dim result_search As String = Nothing
            If items_valuescbCdb.TryGetValue(dgvCashDisburse.Columns(col).HeaderText, result_search) Then
                dgvCashDisburse(col, dgvCashDisburse.RowCount - 1).Value = result_search
            Else
                dgvCashDisburse(col, dgvCashDisburse.RowCount - 1).Value = 0
            End If
        Next
        items_valuescbCdb.Clear()
        creditTotal_Cdr = 0
        debitTotal_Cdr = 0
        lblCreditTotal_Cdb.Text = creditTotal_Cdr
        lbldebitTotal_Cdb.Text = debitTotal_Cdr
    End Sub
    Private Sub btnSubmit_Cdb_Click(sender As Object, e As EventArgs) Handles btnSubmit_Cdb.Click
        If Not String.IsNullOrWhiteSpace(txtboxSupplierN_Cdb.Text) And Not String.IsNullOrWhiteSpace(txtboxRef_Cdb.Text) Then
            If creditTotal_Cdr = debitTotal_Cdr Then
                If dgvCashDisburse.RowCount > 0 Then
                    If Not dgvCashDisburse(2, dgvCashDisburse.RowCount - 1).Value = "Total:" Then
                        insertrow_Cdb()
                    Else
                        Dim msgbox2 As Message_Box = New Message_Box(0, "Cannot add row with a total")
                    End If
                Else
                    insertrow_Cdb()
                End If
            Else
                Dim msgbox3 As Message_Box = New Message_Box(0, "Credit and Debit must be equal")
            End If
        Else
            Dim msgbox4 As Message_Box = New Message_Box(0, "Please complete information")
        End If
    End Sub

    Private Sub btnClear_Cdb_Click(sender As Object, e As EventArgs) Handles btnClear_Cdb.Click
        col_notsortable()
        dgvCashDisburse.Rows.Clear()
        creditTotal_Cdr = 0
        debitTotal_Cdr = 0
        lblCreditTotal_Cdb.Text = creditTotal_Cdr
        lbldebitTotal_Cdb.Text = debitTotal_Cdr
        selected_columnCdr = ""
        txtboxSupplierN_Cdb.Text = ""
        cbSelectColumn_Cbr.SelectedIndex = cbSelectColumn_Cbr.FindStringExact("")
        txtboxValue_Cdb.Text = ""
        txtboxCash_Cdb.Text = ""
        txtboxRef_Cdb.Text = ""
        dtpDate_CDB.Value = Today
    End Sub

    Private Sub btnExport_Cdb_Click(sender As Object, e As EventArgs) Handles btnExport_Cdb.Click
        If dgvCashDisburse.RowCount <> 0 Then
            Dim filename As String
            filename = Microsoft.VisualBasic.DateAndTime.MonthName(Today.Month).ToString & " " & Microsoft.VisualBasic.DateAndTime.Year(Today).ToString & " " & "Cash Disbursement Book"
            Using sfd As SaveFileDialog = New SaveFileDialog With {.Filter = "excel workbook|*.xlsx", .Title = "Export", .FileName = filename}
                Try
                    If sfd.ShowDialog() = DialogResult.OK Then
                        Dim exportxl As ExportExcel = New ExportExcel(dgvCashDisburse.RowCount, dgvCashDisburse.ColumnCount)
                        exportxl.col_headerTitle("Date", "Supplier Description", "Reference", "Credit Cash", "Debit Inventory")
                        For col As Integer = 5 To dgvCashDisburse.ColumnCount - 1
                            exportxl.add_headerTitle(dgvCashDisburse.Columns(col).HeaderText)
                        Next
                        For row As Integer = 0 To dgvCashDisburse.RowCount - 1
                            For col As Integer = 0 To dgvCashDisburse.ColumnCount - 1
                                If (String.IsNullOrWhiteSpace(dgvCashDisburse(col, row).Value) And Not dgvCashDisburse(2, row).Value = "Total:") Then
                                    exportxl.addrow_Items(0)
                                ElseIf dgvCashDisburse(2, row).Value = "Total:" And col < 2 Then
                                    exportxl.addrow_Items("")
                                Else
                                    exportxl.addrow_Items(dgvCashDisburse(col, row).Value.ToString)
                                End If
                            Next
                        Next
                        exportxl.export(sfd.FileName.ToString)
                        col_notsortable()
                        Dim msgbox1 As Message_Box = New Message_Box(1, "Successfully Exported")
                    End If
                Catch ex As Exception
                    Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message)
                End Try
            End Using
        Else
            Dim msgbox3 As Message_Box = New Message_Box(0, "The table is empty")
        End If
    End Sub

    Private Sub btnImport_Cdb_Click(sender As Object, e As EventArgs) Handles btnImport_Cdb.Click
        Using ofd As OpenFileDialog = New OpenFileDialog With {.Filter = "Excel Workbook|*.xlsx", .Title = "Import"}
            Try
                If ofd.ShowDialog() = DialogResult.OK Then
                    If Not (ofd.FileName.ToString.Contains("Cash Disbursement Book")) Then
                        Dim msgbox1 As Message_Box = New Message_Box(0, "Please import only Cash Disbursement Book")
                        Return
                    End If
                    Dim importxl As ImportExcel = New ImportExcel()
                    Dim fileAddress = New FileInfo(ofd.FileName)
                    Dim dataset As System.Data.DataSet

                    dgvCashDisburse.Rows.Clear()
                    dgvCashDisburse.Columns.Clear()
                    dataset = importxl.import(fileAddress.ToString)
                    For i As Integer = 0 To items_comboboxCdb.Count
                        If Not items_comboboxCdb.Count = 2 Then
                            items_comboboxCdb.Pop()
                        End If
                    Next
                    Dim Index As Integer = 1
                    For Each col In dataset.Tables(0).Columns
                        dgvCashDisburse.Columns.Add("Column" & Index, col.ToString)
                        dgvCashDisburse.Columns(Index - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        If Index >= 6 And Not items_comboboxCdb.Contains(col.ToString) Then
                            items_comboboxCdb.Push(col.ToString)
                        End If
                        Index += 1
                    Next
                    For row As Integer = 0 To dataset.Tables(0).Rows.Count - 1
                        dgvCashDisburse.Rows.Add()
                        For col As Integer = 0 To dataset.Tables(0).Columns.Count - 1
                            Dim str_val As String = "Total:"
                            If Not IsDBNull(dataset.Tables(0).Rows(row).Item(col)) Then
                                str_val = dataset.Tables(0).Rows(row).Item(col)
                            End If
                            If row = dataset.Tables(0).Rows.Count - 1 And col < 2 And IsDBNull(dataset.Tables(0).Rows(dataset.Tables(0).Rows.Count - 1).Item(2)) Then
                                str_val = ""
                            End If
                            dgvCashDisburse(col, row).Value = str_val
                        Next
                    Next
                    cbSelectColumn_Cbr.Items.Clear()
                    col_notsortable()
                    For Each item In items_comboboxCdb
                        cbSelectColumn_Cbr.Items.Add(item)
                    Next
                End If
            Catch ex As Exception
                Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message.ToString)
            End Try
        End Using
    End Sub
    Private Sub btnTotal_Cdb_Click(sender As Object, e As EventArgs) Handles btnTotal_Cdb.Click
        If dgvCashDisburse.RowCount <> 0 Then
            If Not dgvCashDisburse(2, dgvCashDisburse.RowCount - 1).Value = "Total:" Then
                Dim total As Double
                dgvCashDisburse.Rows.Add()
                dgvCashDisburse(2, dgvCashDisburse.RowCount - 1).Value = "Total:"
                For col As Integer = 3 To dgvCashDisburse.ColumnCount - 1
                    total = 0
                    For row As Integer = 0 To dgvCashDisburse.RowCount - 1
                        total += dgvCashDisburse(col, row).Value
                    Next
                    dgvCashDisburse(col, dgvCashDisburse.RowCount - 1).Value = total
                Next
            Else
                Dim msgbox1 As Message_Box = New Message_Box(0, "The table has total")
            End If
        Else
            Dim msgbox2 As Message_Box = New Message_Box(0, "The table is empty")
        End If
    End Sub
    Private Sub btnDelete_Cdb_Click(sender As Object, e As EventArgs) Handles btnDelete_Cdb.Click
        If Not (dgvCashDisburse.Rows.Count = 0) Then
            dgvCashDisburse.Rows.RemoveAt(dgvCashDisburse.Rows.Count - 1)
        Else
            Dim msgbox1 As Message_Box = New Message_Box(0, "You cannot delete a row in an empty table.")
        End If
    End Sub
    'cash disbursement
    'journal
    Private Sub lnklblSalescredit_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnklblSalescredit.LinkClicked
        Dim vsj As ViewSampleJournal = New ViewSampleJournal()
        vsj.set_Title("Sales on Credit")
        vsj.set_Image()
        vsj.Show()
    End Sub

    Private Sub lnklblPurchasecredit_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnklblPurchasecredit.LinkClicked
        Dim vsj As ViewSampleJournal = New ViewSampleJournal()
        vsj.set_Title("Purchase on Credit")
        vsj.set_Image()
        vsj.Show()
    End Sub
    Private Sub lnklblInventory_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnklblInventory.LinkClicked
        Dim vsj As ViewSampleJournal = New ViewSampleJournal()
        vsj.set_Title("Recording quarterly or year end inventory")
        vsj.set_Image()
        vsj.Show()
    End Sub
    Private Sub lnklblDepreciation_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnklblDepreciation.LinkClicked
        Dim vsj As ViewSampleJournal = New ViewSampleJournal()
        vsj.set_Title("Recording of depreciation on assets")
        vsj.set_Image()
        vsj.Show()
    End Sub
    Private Sub btnSubmit_Gj_Click(sender As Object, e As EventArgs) Handles btnSubmit_Gj.Click
        Dim data_name() As String = {"Date", "Description", "Debit", "Credit", "Value"}
        Dim data() As String = {dtpDate_Gj.Text, txtboxDescription_Gj.Text, cbDebit_Gj.Text, cbCredit_Gj.Text, txtboxValue_Gj.Text}
        Dim count_empty As Integer = 0
        Dim message As String = ""
        Dim value As Double
        For index As Integer = 0 To 4
            If (String.IsNullOrWhiteSpace(data(index))) Then
                message = message + data_name(index) & "|"
                count_empty += 1
            End If
        Next
        If (count_empty >= 1) Then
            Dim msgbox1 As Message_Box = New Message_Box(0, message & " is/are empty")
        ElseIf Not Double.TryParse(data(4), value) Then
            Dim msgbox2 As Message_Box = New Message_Box(0, "Please enter number only")
        ElseIf Double.TryParse(data(4), value) Then
            dgvJournal.Rows.Add(3)
            Dim rows_c As Integer = dgvJournal.Rows.Count - 1
            dgvJournal(0, rows_c - 2).Value = dgvJournal.Rows.Count / 3
            dgvJournal(1, rows_c - 2).Value = dtpDate_Gj.Text
            dgvJournal(2, rows_c - 2).Value = cbDebit_Gj.Text
            dgvJournal(2, rows_c - 1).Value = cbCredit_Gj.Text
            dgvJournal(2, rows_c).Value = txtboxDescription_Gj.Text
            dgvJournal(3, rows_c - 2).Value = txtboxValue_Gj.Text
            dgvJournal(4, rows_c - 1).Value = txtboxValue_Gj.Text
            dtpDate_Gj.Text = dtpDate_Gj.Value.AddDays(1)
            txtboxValue_Gj.Text = 0
            txtboxDescription_Gj.Text = ""
        End If
    End Sub

    Private Sub btnClear_Gj_Click(sender As Object, e As EventArgs) Handles btnClear_Gj.Click
        dgvJournal.Rows.Clear()
        dtpDate_Gj.Text = Today
        cbCredit_Gj.Text = Nothing
        cbDebit_Gj.Text = Nothing
        txtboxDescription_Gj.Text = ""
        txtboxValue_Gj.Text = ""
    End Sub

    Private Sub btnDeleteRow_Gj_Click(sender As Object, e As EventArgs) Handles btnDeleteRow_Gj.Click
        If Not (dgvJournal.Rows.Count = 0) Then
            For index As Integer = 3 To 1 Step -1
                dgvJournal.Rows.RemoveAt(dgvJournal.Rows.Count - index)
            Next
        Else
            Dim msgbox1 As Message_Box = New Message_Box(0, "You cannot delete a row in an empty table.")
        End If
    End Sub

    Private Sub btnExport_Gj_Click(sender As Object, e As EventArgs) Handles btnExport_Gj.Click
        If dgvJournal.RowCount <> 0 Then
            Dim filename As String
            filename = Microsoft.VisualBasic.DateAndTime.MonthName(Today.Month).ToString & " " & Microsoft.VisualBasic.DateAndTime.Year(Today).ToString & " " & "General Journal Book"
            Using sfd As SaveFileDialog = New SaveFileDialog With {.Filter = "excel workbook|*.xlsx", .Title = "Export", .FileName = filename}
                Try
                    If sfd.ShowDialog() = DialogResult.OK Then
                        Dim exportxl As ExportExcel = New ExportExcel(dgvJournal.RowCount, dgvJournal.ColumnCount)
                        exportxl.col_headerTitle("No.", "Date", "Account Title", "Debit", "Credit")
                        For row As Integer = 0 To dgvJournal.RowCount - 1
                            For col As Integer = 0 To dgvJournal.ColumnCount - 1
                                If Not IsDBNull(dgvJournal(col, row).Value) Then
                                    If (String.IsNullOrWhiteSpace(dgvJournal(col, row).Value)) Then
                                        exportxl.addrow_Items("")
                                    Else
                                        exportxl.addrow_Items(dgvJournal(col, row).Value.ToString)
                                    End If
                                Else
                                    exportxl.addrow_Items("")
                                End If
                            Next
                        Next
                        exportxl.export(sfd.FileName.ToString)
                        Dim msgbox1 As Message_Box = New Message_Box(1, "Successfully Exported")
                    End If
                Catch ex As Exception
                    Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message.ToString)
                End Try
            End Using
        Else
            Dim msgbox3 As Message_Box = New Message_Box(0, "The table is empty")
        End If
    End Sub

    Private Sub btnImport_Gj_Click(sender As Object, e As EventArgs) Handles btnImport_Gj.Click
        Using ofd As OpenFileDialog = New OpenFileDialog With {.Filter = "Excel Workbook|*.xlsx", .Title = "Import"}
            Try
                If ofd.ShowDialog() = DialogResult.OK Then
                    If Not (ofd.FileName.ToString.Contains("General Journal Book")) Then
                        Dim msgbox1 As Message_Box = New Message_Box(0, "Please import only General Journal Book")
                        Return
                    End If
                    Dim importxl As ImportExcel = New ImportExcel()
                    Dim fileAddress = New FileInfo(ofd.FileName)
                    Dim dataset As System.Data.DataSet

                    dgvJournal.Rows.Clear()
                    dataset = importxl.import(fileAddress.ToString)
                    For row As Integer = 0 To dataset.Tables(0).Rows.Count - 1
                        Dim date_str As String = ""
                        If Not (IsDBNull(dataset.Tables(0).Rows(row).Item(1))) Then
                            date_str = dataset.Tables(0).Rows(row).Item(1)
                        End If
                        dgvJournal.Rows.Add(dataset.Tables(0).Rows(row).Item(0), date_str, dataset.Tables(0).Rows(row).Item(2), dataset.Tables(0).Rows(row).Item(3), dataset.Tables(0).Rows(row).Item(4))
                    Next
                End If
            Catch ex As Exception
                Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message.ToString)
            End Try
        End Using
    End Sub
    'journal
    'ledger
    Private Sub cbAccountT_L_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbAccountT_L.SelectedIndexChanged
        lblTitle_Ledger.Text = cbAccountT_L.Text
    End Sub
    Private Sub cbAccountT_L_TextUpdate(sender As Object, e As EventArgs) Handles cbAccountT_L.TextUpdate
        If String.IsNullOrWhiteSpace(cbAccountT_L.Text) Then
            lblTitle_Ledger.Text = "Choose account title"
        Else
            lblTitle_Ledger.Text = cbAccountT_L.Text
        End If
    End Sub
    Private Sub btnSubmit_Ledger_Click(sender As Object, e As EventArgs) Handles btnSubmit_Ledger.Click
        If dgvLedger.Rows.Count >= 1 Then
            If Not (IsDBNull(dgvLedger(3, dgvLedger.Rows.Count - 1).Value)) Then
                If dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Total for the quarter:" Then
                    Dim msgbox1 As Message_Box = New Message_Box(0, "Can't submit if there is a total")
                    Return
                End If
            End If
        End If
        If Not isCreditEmpty() Or Not isDebitEmpty() Then
            Dim debit As Double = 0
            Dim credit As Double = 0
            If Not String.IsNullOrEmpty(txtboxDebit_L.Text) Then
                If Not Double.TryParse(txtboxDebit_L.Text, debit) Then
                    Dim msgbox1 As Message_Box = New Message_Box(0, "Debit accepts only number")
                    Return
                End If
            End If
            If Not String.IsNullOrEmpty(txtboxCredit_L.Text) Then
                If Not Double.TryParse(txtboxCredit_L.Text, credit) Then
                    Dim msgbox2 As Message_Box = New Message_Box(0, "Credit accepts only number")
                    Return
                End If
            End If
            dgvLedger.Rows.Add()
            Dim data() As String = {dgvLedger.Rows.Count, cbMonth_Ledger.Text & "-" & cbYear_Ledger.Text, txtboxDebitDesc_L.Text, txtboxRefD_Ledger.Text, txtboxDebit_L.Text, txtboxCreditDesc_L.Text, txtboxRefC_Ledger.Text, txtboxCredit_L.Text}
            Dim start_inc As Integer = 0
            dgvLedger(0, dgvLedger.Rows.Count - 1).Value = data(0)
            dgvLedger(1, dgvLedger.Rows.Count - 1).Value = data(1)
            If String.IsNullOrWhiteSpace(data(2)) Or String.IsNullOrWhiteSpace(data(4)) Then
                start_inc = 5
            End If
            For index As Integer = start_inc To 4
                dgvLedger(index, dgvLedger.Rows.Count - 1).Value = data(index)
            Next
            If String.IsNullOrWhiteSpace(data(5)) Or String.IsNullOrWhiteSpace(data(7)) Then
                start_inc = 8
            End If
            For index As Integer = start_inc To 7
                dgvLedger(index, dgvLedger.Rows.Count - 1).Value = data(index)
            Next
            txtboxDebitDesc_L.Clear()
            txtboxDebit_L.Clear()
            txtboxRefD_Ledger.Clear()
            txtboxCreditDesc_L.Clear()
            txtboxCredit_L.Clear()
            txtboxRefD_Ledger.Clear()
            txtboxRefC_Ledger.Clear()
        Else
            Dim check_empty() As String = {txtboxDebitDesc_L.Text, txtboxDebit_L.Text, txtboxCreditDesc_L.Text, txtboxCredit_L.Text}
            Dim data_name() As String = {"Debit Description", "Debit Value", "Credit Description", "Credit Value"}
            Dim empty_list As String = ""
            For index As Integer = 0 To check_empty.Count - 1
                If String.IsNullOrWhiteSpace(check_empty(index)) Then
                    empty_list = empty_list + data_name(index) + "|"
                End If
            Next
            Dim msgbox3 As Message_Box = New Message_Box(0, empty_list & " is/are empty")
        End If
    End Sub
    Private Function isDebitEmpty()
        If String.IsNullOrWhiteSpace(txtboxDebitDesc_L.Text) Or String.IsNullOrWhiteSpace(txtboxDebit_L.Text) Then
            Return True
        End If
        Return False
    End Function
    Private Function isCreditEmpty()
        If String.IsNullOrWhiteSpace(txtboxCreditDesc_L.Text) Or String.IsNullOrWhiteSpace(txtboxCredit_L.Text) Then
            Return True
        End If
        Return False
    End Function

    Private Sub btnDel_Ledger_Click(sender As Object, e As EventArgs) Handles btnDel_Ledger.Click
        If Not (dgvLedger.Rows.Count = 0) Then
            If String.IsNullOrWhiteSpace(dgvLedger(1, dgvLedger.Rows.Count - 1).Value) Then
                If dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Total for the quarter:" Then
                    dgvLedger.Rows.RemoveAt(dgvLedger.Rows.Count - 1)
                    dgvLedger.Rows.RemoveAt(dgvLedger.Rows.Count - 1)
                    Return
                End If
            End If
            dgvLedger.Rows.RemoveAt(dgvLedger.Rows.Count - 1)
        Else
            Dim msgbox1 As Message_Box = New Message_Box(0, "You cannot delete a row in an empty table.")
        End If
    End Sub

    Private Sub btnClead_Ledger_Click(sender As Object, e As EventArgs) Handles btnClear_Ledger.Click
        cbMonth_Ledger.SelectedIndex = Microsoft.VisualBasic.Month(Today.Date) - 1
        cbYear_Ledger.Text = Microsoft.VisualBasic.DateAndTime.Year(Today).ToString
        txtboxDebit_L.Clear()
        txtboxCredit_L.Clear()
        txtboxCreditDesc_L.Clear()
        txtboxDebitDesc_L.Clear()
        txtboxRefC_Ledger.Clear()
        txtboxRefD_Ledger.Clear()
        cbAccountT_L.Text = ""
        dgvLedger.Rows.Clear()
        lblTitle_Ledger.Text = "Choose account title"
    End Sub

    Private Sub btnExport_Ledger_Click(sender As Object, e As EventArgs) Handles btnExport_Ledger.Click
        If String.IsNullOrEmpty(cbAccountT_L.Text) Then
            Dim msgbox1 As Message_Box = New Message_Box(0, "Please choose account title")
            Return
        End If
        If dgvLedger.RowCount <> 0 Then
            Dim filename As String
            filename = Microsoft.VisualBasic.DateAndTime.MonthName(Today.Month).ToString & " " & Microsoft.VisualBasic.DateAndTime.Year(Today).ToString & " " & cbAccountT_L.Text & " " & "Ledger"
            Using sfd As SaveFileDialog = New SaveFileDialog With {.Filter = "excel workbook|*.xlsx", .Title = "Export", .FileName = filename}
                Try
                    If sfd.ShowDialog() = DialogResult.OK Then
                        Dim exportxl As ExportExcel = New ExportExcel(dgvLedger.RowCount, dgvLedger.ColumnCount)
                        exportxl.col_headerTitle("No.", "Date", "Debit Description", "Debit Ref", "Debit", "Credit Description", "Credit Ref", "Credit")
                        For row As Integer = 0 To dgvLedger.RowCount - 1
                            For col As Integer = 0 To dgvLedger.ColumnCount - 1
                                If Not IsDBNull(dgvLedger(col, row).Value) Then
                                    If (String.IsNullOrWhiteSpace(dgvLedger(col, row).Value)) Then
                                        exportxl.addrow_Items("")
                                    Else
                                        exportxl.addrow_Items(dgvLedger(col, row).Value)
                                    End If
                                Else
                                    exportxl.addrow_Items("")
                                End If
                            Next
                        Next
                        exportxl.export_customDT(sfd.FileName.ToString)
                        Dim msgbox2 As Message_Box = New Message_Box(1, "Successfully Exported")
                    End If
                Catch ex As Exception
                    Dim msgbox3 As Message_Box = New Message_Box(0, ex.Message.ToString)
                End Try
            End Using
        Else
            Dim msgbox4 As Message_Box = New Message_Box(0, "The table is empty")
        End If
    End Sub

    Private Sub btnImport_Ledger_Click(sender As Object, e As EventArgs) Handles btnImport_Ledger.Click
        Using ofd As OpenFileDialog = New OpenFileDialog With {.Filter = "Excel Workbook|*.xlsx", .Title = "Import"}
            Try
                If ofd.ShowDialog() = DialogResult.OK Then
                    If Not (ofd.FileName.ToString.Contains("Ledger")) Then
                        Dim msgbox1 As Message_Box = New Message_Box(0, "Please import only Ledger Book")
                        Return
                    End If
                    Dim importxl As ImportExcel = New ImportExcel()
                    Dim fileAddress = New FileInfo(ofd.FileName)
                    Dim dataset As System.Data.DataSet

                    dgvLedger.Rows.Clear()
                    cbAccountT_L.Text = " "
                    cbAccountT_L.SelectedText = return_accT(ofd.FileName.ToString)
                    dataset = importxl.import(fileAddress.ToString)
                    For row As Integer = 0 To dataset.Tables(0).Rows.Count - 1
                        If IsDBNull(dataset.Tables(0).Rows(row).Item(0)) Then
                            Exit For
                        End If
                        Dim date_str As String = " "
                        If Not IsDBNull(dataset.Tables(0).Rows(row).Item(1)) And Not IsDBNull(dataset.Tables(0).Rows(row).Item(0)) Then
                            Dim date_data As DateTime = Convert.ToDateTime(dataset.Tables(0).Rows(row).Item(1))
                            date_str = date_data.ToString("MMM-yyyy")
                        End If
                        dgvLedger.Rows.Add(dataset.Tables(0).Rows(row).Item(0), date_str, dataset.Tables(0).Rows(row).Item(2), dataset.Tables(0).Rows(row).Item(3), dataset.Tables(0).Rows(row).Item(4), dataset.Tables(0).Rows(row).Item(5), dataset.Tables(0).Rows(row).Item(6), dataset.Tables(0).Rows(row).Item(7))
                    Next
                    If IsDBNull(dataset.Tables(0).Rows(dgvLedger.Rows.Count - 1).Item(1)) And Not IsDBNull(dataset.Tables(0).Rows(dgvLedger.Rows.Count - 1).Item(0)) Then
                        dgvLedger(3, dgvLedger.Rows.Count - 2).Value = "Subtotal:"
                        dgvLedger(6, dgvLedger.Rows.Count - 2).Value = "Subtotal:"
                        dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Total for the quarter:"
                    End If
                End If
            Catch ex As Exception
                Dim msgbox2 As Message_Box = New Message_Box(0, ex.Message.ToString)
            End Try
        End Using
    End Sub

    Private Sub btnTotal_Ledger_Click(sender As Object, e As EventArgs) Handles btnTotal_Ledger.Click
        If dgvLedger.Rows.Count = 0 Then
            Dim msgbox1 As Message_Box = New Message_Box(0, "There are no rows to total")
            Return
        End If
        If Not (IsDBNull(dgvLedger(3, dgvLedger.Rows.Count - 1).Value)) Then
            If dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Total for the quarter:" Then
                Dim msgbox2 As Message_Box = New Message_Box(0, "The table has total")
                Return
            End If
        End If
        Dim total_Debit As Double
        Dim total_Credit As Double
        For row As Integer = 0 To dgvLedger.Rows.Count - 1
            Dim val As Double
            If IsDBNull(dgvLedger(4, row).Value) Then
                Continue For
            End If
            If Double.TryParse(dgvLedger(4, row).Value, val) Then
                total_Debit += val
            End If
        Next
        For row As Integer = 0 To dgvLedger.Rows.Count - 1
            Dim val As Double
            If IsDBNull(dgvLedger(7, row).Value) Then
                Continue For
            End If
            If Double.TryParse(dgvLedger(7, row).Value, val) Then
                total_Credit += val
            End If
        Next
        dgvLedger.Rows.Add()
        dgvLedger(0, dgvLedger.Rows.Count - 1).Value = dgvLedger.Rows.Count
        dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Subtotal:"
        dgvLedger(4, dgvLedger.Rows.Count - 1).Value = total_Debit
        dgvLedger(6, dgvLedger.Rows.Count - 1).Value = "Subtotal:"
        dgvLedger(7, dgvLedger.Rows.Count - 1).Value = total_Credit
        dgvLedger.Rows.Add()
        dgvLedger(0, dgvLedger.Rows.Count - 1).Value = dgvLedger.Rows.Count
        dgvLedger(3, dgvLedger.Rows.Count - 1).Value = "Total for the quarter:"
        If (total_Debit > total_Credit) Then
            dgvLedger(4, dgvLedger.Rows.Count - 1).Value = total_Debit - total_Credit
            Return
        End If
        dgvLedger(4, dgvLedger.Rows.Count - 1).Value = total_Credit - total_Debit
    End Sub

    Private Function return_accT(ByVal filename As String)
        For index As Integer = 0 To cbAccountT_L.Items.Count - 1
            If (filename.Contains(cbAccountT_L.Items(index))) Then
                Return cbAccountT_L.Items(index)
            End If
        Next
        Return " "
    End Function

    'ledger
End Class
