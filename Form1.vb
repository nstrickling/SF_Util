Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports System.Collections.Generic
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmMain
    Public XLSDictionary As New Dictionary(Of String, String) 'Data dictionary including the operator name and ID (for lookup)
    Public Function GetPACPData(ByVal sConnection As String) As Boolean
        '##################################################################################################################################################
        '#      Function:       GetPACPData()
        '#      Description:    Get the PACP data
        '#      Version:        1.0
        '#      Last changed:   23.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Dim sSQL As String

        sSQL = "SELECT Inspections.InspectionID, Inspections.Pipe_Segment_Reference, Inspections.Upstream_MH, Inspections.Downstream_MH, Inspections.Street, Inspections.City, Inspections.Location_Details FROM Inspections ORDER BY Inspections.InspectionID"

        Using connection As New OleDbConnection(sConnection)
            Dim command As New OleDbCommand(sSQL)
            Dim dAdapter As OleDbDataAdapter
            dAdapter = New OleDbDataAdapter()
            dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

            command.Connection = connection

            connection.Open()
            command.ExecuteNonQuery()
            Dim dSet As DataSet = New DataSet()
            Dim dTable As DataTable = New DataTable()
            dAdapter.Fill(dSet, "Inspections")
            dAdapter.Fill(dTable)

            dTable.Columns.Add("Status", Type.GetType("System.String"))

            Me.DataGridView1.DataSource = dTable

            Me.DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

            dAdapter.Dispose()
            command.Dispose()
            dTable.Dispose()
            dSet.Dispose()
            connection.Close() 'Closes the connection

            Me.ProgressBar1.Maximum = DataGridView1.Rows.Count

            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                If row.IsNewRow Then Exit For

                Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
            Next
        End Using
    End Function

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        '##################################################################################################################################################
        '#      Function:       DataGridView1_ColumnHeaderMouseClick()
        '#      Description:    Sorting data grid view (after double-click on column header, recolor cells based on value)
        '#      Version:        1.0
        '#      Last changed:   21.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            If row.IsNewRow Then Exit For
            If row.Cells(7).Value.ToString = "N/A" Then
                row.Cells(7).Style.BackColor = Color.Red
            Else
                row.Cells(7).Style.BackColor = Color.Green
            End If
        Next
    End Sub

    Private Sub ExportMediaLog()
        '##################################################################################################################################################
        '#      Function:       ExportToExcelToolStripMenuItem_Click()
        '#      Description:    Export data to Excel
        '#      Version:        1.0
        '#      Last changed:   21.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Dim sTable(Me.DataGridView2.Rows.Count + Me.DataGridView3.Rows.Count + 1, 3) As String
        Dim i As Integer
        Dim xApp As Excel.Application
        Dim xBook As Excel.Workbook
        Dim xSheet As Excel.Worksheet
        Dim j As Integer

        If Me.DataGridView2.Rows.Count = 0 Then
            MsgBox("No data available", vbInformation, "Excel Export")
            Exit Sub
        End If

        Me.Panel1.Visible = True

        Me.MenuStrip1.Enabled = False

        For Each row As DataGridViewRow In DataGridView2.Rows
            If row.IsNewRow = True Then
                Exit For
            End If

            sTable(i, 0) = row.Cells(0).Value.ToString 'WorkOrder
            sTable(i, 1) = row.Cells(1).Value.ToString 'Maximo Asset Number
            sTable(i, 2) = row.Cells(2).Value.ToString 'Maximo Asset Number

            Select Case row.Cells(3).Style.BackColor 'Coded by
                Case Color.Red
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Red"
                Case Color.Orange
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Orange"
                Case Color.Green
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Green"
                Case Else
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=White"
            End Select

            i = i + 1
        Next
        sTable(i, 0) = "ID"
        sTable(i, 1) = "Media Name (Video)"
        sTable(i, 2) = "Media Path (Video)"
        sTable(i, 3) = "Status (Video)BGCOLOR=White"
        i = i + 1

        For Each row As DataGridViewRow In DataGridView3.Rows
            If row.IsNewRow = True Then
                Exit For
            End If

            sTable(i, 0) = row.Cells(0).Value.ToString 'WorkOrder
            sTable(i, 1) = row.Cells(1).Value.ToString 'Maximo Asset Number
            sTable(i, 2) = row.Cells(2).Value.ToString 'Maximo Asset Number

            Select Case row.Cells(3).Style.BackColor 'Coded by
                Case Color.Red
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Red"
                Case Color.Orange
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Orange"
                Case Color.Green
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=Green"
                Case Else
                    sTable(i, 3) = row.Cells(3).Value.ToString & "BGCOLOR=White"
            End Select

            i = i + 1
        Next

        xApp = CreateObject("Excel.Application")
        'xApp.Visible = True
        xBook = xApp.Workbooks.Add
        xSheet = xBook.ActiveSheet

        xSheet.Cells(1, 1) = "ID"
        xSheet.Cells(1, 2) = "Media Name (Photo)"
        xSheet.Cells(1, 3) = "Media Path (Photo)"
        xSheet.Cells(1, 4) = "Status (Photo)"


        Me.ProgressBar1.Maximum = UBound(sTable)
        For i = 0 To UBound(sTable) - 1
            xSheet.Cells(i + 2, 1) = sTable(i, 0).ToString
            xSheet.Cells(i + 2, 2) = sTable(i, 1).ToString
            xSheet.Cells(i + 2, 3) = sTable(i, 1).ToString
            xSheet.Cells(i + 2, 4) = Mid(sTable(i, 3).ToString, 1, InStr(sTable(i, 3), "BGCOLOR", vbTextCompare) - 1)

            Select Case Mid(sTable(i, 3), InStr(sTable(i, 3), "BGCOLOR", vbTextCompare) + 8, 50)
                Case "Red"
                    xSheet.Cells(i + 2, 4).interior.colorindex = "3"
                Case "Green"
                    xSheet.Cells(i + 2, 4).interior.colorindex = "4"
                Case "Orange"
                    xSheet.Cells(i + 2, 4).interior.colorindex = "46"

            End Select


            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next


        With xSheet.Range(xSheet.Cells(1, 1), xSheet.Cells(UBound(sTable), 4))
            .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        End With

        Dim sExcelFile As String

        sExcelFile = Application.StartupPath & "\Reports\ExcelReports_PACP_MediaLOG_" & Guid.NewGuid.ToString & "_" & DateAndTime.Day(Now) & "_" & DateAndTime.Month(Now) & "_" & DateAndTime.Year(Now) & ".xls"
        xBook.SaveAs(sExcelFile)
        If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then Me.ProgressBar1.Value = Me.ProgressBar1.Maximum
        If MsgBox("Excel export finished. Do you want to open the exported file?", vbYesNo, "Excel Export") = MsgBoxResult.Yes Then
            xBook.Close()
            xSheet = Nothing
            xApp = Nothing
            System.Diagnostics.Process.Start(sExcelFile)
        End If
        Me.Panel1.Visible = False
        Me.Label1.Visible = False

        Me.ProgressBar1.Value = 0

        Me.MenuStrip1.Enabled = True
    End Sub

    Private Sub ExportPACPLog()
        '##################################################################################################################################################
        '#      Function:       ExportToExcelToolStripMenuItem_Click()
        '#      Description:    Export data to Excel
        '#      Version:        1.0
        '#      Last changed:   21.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Dim sTable(Me.DataGridView1.Rows.Count, 8) As String
        Dim i As Integer
        Dim xApp As Excel.Application
        Dim xBook As Excel.Workbook
        Dim xSheet As Excel.Worksheet
        Dim j As Integer

        If Me.DataGridView1.Rows.Count = 0 Then
            MsgBox("No data available", vbInformation, "Excel Export")
            Exit Sub
        End If

        Me.Panel1.Visible = True

        Me.MenuStrip1.Enabled = False

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.IsNewRow = True Then
                Exit For
            End If

            sTable(i, 0) = row.Cells(0).Value.ToString 'WorkOrder
            sTable(i, 1) = row.Cells(1).Value.ToString 'Maximo Asset Number
            sTable(i, 2) = row.Cells(2).Value.ToString 'Date
            sTable(i, 3) = row.Cells(3).Value.ToString 'Surveyed By
            sTable(i, 4) = row.Cells(4).Value.ToString 'Coded By
            sTable(i, 5) = row.Cells(5).Value.ToString 'Coded By
            sTable(i, 6) = row.Cells(6).Value.ToString 'Coded By


            Select Case row.Cells(7).Style.BackColor 'Coded by
                Case Color.Red
                    sTable(i, 7) = row.Cells(7).Value.ToString & "BGCOLOR=Red"
                Case Color.Green
                    sTable(i, 7) = row.Cells(7).Value.ToString & "BGCOLOR=Green"
                Case Else
                    sTable(i, 7) = row.Cells(7).Value.ToString & "BGCOLOR=White"
            End Select



            i = i + 1
        Next

        xApp = CreateObject("Excel.Application")
        'xApp.Visible = True
        xBook = xApp.Workbooks.Add
        xSheet = xBook.ActiveSheet

        xSheet.Cells(1, 1) = "Inspection_ID#"
        xSheet.Cells(1, 2) = "Pipe_Segment_Reference"
        xSheet.Cells(1, 3) = "Upstream_MH"
        xSheet.Cells(1, 4) = "Downstream_MH"
        xSheet.Cells(1, 5) = "Street"
        xSheet.Cells(1, 6) = "City"
        xSheet.Cells(1, 7) = "Location_Details"
        xSheet.Cells(1, 8) = "Status"


        Me.ProgressBar1.Maximum = UBound(sTable)
        For i = 0 To UBound(sTable) - 1
            xSheet.Cells(i + 2, 1) = sTable(i, 0).ToString
            xSheet.Cells(i + 2, 2) = sTable(i, 1).ToString
            xSheet.Cells(i + 2, 3) = sTable(i, 2).ToString
            xSheet.Cells(i + 2, 4) = sTable(i, 3).ToString
            xSheet.Cells(i + 2, 5) = sTable(i, 4).ToString
            xSheet.Cells(i + 2, 6) = sTable(i, 5).ToString
            xSheet.Cells(i + 2, 7) = sTable(i, 6).ToString
            xSheet.Cells(i + 2, 8) = Mid(sTable(i, 7).ToString, 1, InStr(sTable(i, 7), "BGCOLOR", vbTextCompare) - 1)

            Select Case Mid(sTable(i, 7), InStr(sTable(i, 7), "BGCOLOR", vbTextCompare) + 8, 50)
                Case "Red"
                    xSheet.Cells(i + 2, 8).interior.colorindex = "3"
                Case "Green"
                    xSheet.Cells(i + 2, 8).interior.colorindex = "4"
            End Select


            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next


        With xSheet.Range(xSheet.Cells(1, 1), xSheet.Cells(UBound(sTable), 9))
            .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        End With

        Dim sExcelFile As String

        sExcelFile = Application.StartupPath & "\Reports\ExcelReports_PACP_Update_" & Guid.NewGuid.ToString & "_" & DateAndTime.Day(Now) & "_" & DateAndTime.Month(Now) & "_" & DateAndTime.Year(Now) & ".xls"
        xBook.SaveAs(sExcelFile)
        If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then Me.ProgressBar1.Value = Me.ProgressBar1.Maximum
        If MsgBox("Excel export finished. Do you want to open the exported file?", vbYesNo, "Excel Export") = MsgBoxResult.Yes Then
            xBook.Close()
            xSheet = Nothing
            xApp = Nothing
            System.Diagnostics.Process.Start(sExcelFile)
        End If
        Me.Panel1.Visible = False
        Me.Label1.Visible = False

        Me.ProgressBar1.Value = 0

        Me.MenuStrip1.Enabled = True
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        If Me.TabControl1.SelectedTab.Name.ToString = "TabPage1" Then
            ExportPACPLog()
        End If

        If Me.TabControl1.SelectedTab.Name.ToString = "TabPage2" Then
            ExportMediaLog()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub OpenPACPDatabaseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenPACPDatabaseToolStripMenuItem.Click
        '##################################################################################################################################################
        '#      Function:       OpenPACPDatabaseToolStripMenuItem_Click()
        '#      Description:    Opens the file browse dialogue
        '#      Version:        1.0
        '#      Last changed:   22.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Dim sConnection As String
        Me.OpenFileDialog1.Title = "Select PACPdatabase"
        Me.OpenFileDialog1.Filter = "All files (*.MDB)|*.MDB|All files (*.MDB)|*.MDB"

        Me.ProgressBar1.Value = 0
        Me.OpenFileDialog1.RestoreDirectory = True

        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Panel1.Visible = True 'Panel1 shows the WinCan Logo and a label with a user information like "DB check in progress, please be patient"
            Me.Panel1.Refresh()
            Me.lblFilePath.Text = Me.OpenFileDialog1.FileName
            Me.DataGridView1.DataSource = Nothing
            Me.DataGridView1.Rows.Clear()
            Me.DataGridView1.Columns.Clear()
            Me.DataGridView1.Refresh()

            sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"
            Me.Panel1.Visible = True
            Me.lblStatus.Text = "Loading data. Please be patient ..."
            GetPACPData(sConnection)
            If Me.lblMediaPath.Text <> "-" Then
                Me.cmdCopyMedia.Enabled = True
            Else
                Me.cmdCopyMedia.Enabled = False
            End If

            Me.ProgressBar1.Value = 0
            Me.lblStatus.Text = "Loading media files. Please be patient ..."
            LoadMedia()

            Me.Panel1.Visible = False 'Hide Panel1 again
            If Me.lblXLSLookUp.Text <> "-" Then
                Me.cmdUpdate.Enabled = True
            End If
            Me.ProgressBar1.Value = 0
        End If


    End Sub

    Private Sub LoadMedia()
        Dim sConnection As String
        Dim sClipPath As String
        Dim sTMP() As String

        Me.lblStatus.Text = "Loading photo files. Please be patient ..."
        'PHOTO
        Me.DataGridView2.DataSource = Nothing
        Me.DataGridView2.Rows.Clear()
        Me.DataGridView2.Columns.Clear()
        Me.DataGridView2.Refresh()

        sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"

        Dim sSQL As String

        sSQL = "SELECT MediaCondID, Image_Reference, Image_Path from Media_Conditions UNION  SELECT MediaCondID, Image_Reference, Image_Path from LACP_Media_Conditions"

        Using connection As New OleDbConnection(sConnection)
            Dim command As New OleDbCommand(sSQL)
            Dim dAdapter As OleDbDataAdapter
            dAdapter = New OleDbDataAdapter()
            dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

            command.Connection = connection

            connection.Open()
            command.ExecuteNonQuery()
            Dim dSet As DataSet = New DataSet()
            Dim dTable As DataTable = New DataTable()
            dAdapter.Fill(dSet, "Media_Conditions")
            dAdapter.Fill(dTable)

            dTable.Columns.Add("Status", Type.GetType("System.String"))

            Me.DataGridView2.DataSource = dTable

            Me.DataGridView2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

            dAdapter.Dispose()
            command.Dispose()
            dTable.Dispose()
            dSet.Dispose()
            connection.Close() 'Closes the connection

            Me.ProgressBar1.Maximum = DataGridView1.Rows.Count

            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                If row.IsNewRow Then Exit For

                Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
            Next
        End Using


        'Video

        Me.lblStatus.Text = "Loading video files. Please be patient ..."
        Me.DataGridView3.DataSource = Nothing
        Me.DataGridView3.Rows.Clear()
        Me.DataGridView3.Columns.Clear()
        Me.DataGridView3.Refresh()

        sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"



        sSQL = "SELECT MediaID, Video_Name, Video_Location from Media_Inspections WHERE Video_Name NOT LIKE '%.ipf' UNION SELECT MediaID, Video_Name, Video_Location from LACP_Media_Inspections WHERE Video_Name NOT LIKE '%.ipf'"

        Using connection As New OleDbConnection(sConnection)
            Dim command As New OleDbCommand(sSQL)
            Dim dAdapter As OleDbDataAdapter
            dAdapter = New OleDbDataAdapter()
            dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

            command.Connection = connection

            connection.Open()
            command.ExecuteNonQuery()
            Dim dSet As DataSet = New DataSet()
            Dim dTable As DataTable = New DataTable()
            dAdapter.Fill(dSet, "Media_Conditions")
            dAdapter.Fill(dTable)

            dTable.Columns.Add("Status", Type.GetType("System.String"))

            Me.DataGridView3.DataSource = dTable

            Me.DataGridView3.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

            dAdapter.Dispose()
            command.Dispose()
            dTable.Dispose()
            dSet.Dispose()
            connection.Close() 'Closes the connection

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = DataGridView1.Rows.Count

            For Each row As DataGridViewRow In Me.DataGridView3.Rows
                If row.IsNewRow Then Exit For
                sClipPath = row.Cells(2).Value
                Exit For
            Next
        End Using

        If Strings.Left(sClipPath, 1) = "\" Or Mid(sClipPath, 2, 1) = ":" Then
            sTMP = Split(sClipPath, "\")
            modInivb.sNewVidPath = sTMP(UBound(sTMP)) & "\"
            'sClipPath = Mid(sClipPath, 2, Len(sClipPath))
            Me.lblVideoSourcePath.Text = sClipPath
            modInivb.sStaticPath = True
        Else
            modInivb.sNewVidPath = ""
            modInivb.sStaticPath = False
        End If

    End Sub

    Public Sub GetXLSData(ByVal sFile As String)
        '##################################################################################################################################################
        '#      Function:       GetXLSDATA()
        '#      Description:    Get xls Data and save in DataDictionary
        '#      Version:        1.0
        '#      Last changed:   23.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection _
        ("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & sFile & "';Extended Properties=Excel 8.0;")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter _
            ("select * from [Sheet1$]", MyConnection)
        MyCommand.TableMappings.Add("Table", "TestTable")
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)

        MyConnection.Close()

        Me.XLSDictionary.Clear()

        Me.ProgressBar1.Maximum = DtSet.Tables(0).Rows.Count
        For Each Row As DataRow In DtSet.Tables(0).Rows
            If XLSDictionary.ContainsKey(Row(0).ToString) = False Then
                XLSDictionary.Add(Row(0).ToString, Row(1).ToString)
                If Me.ProgressBar1.Value > Me.ProgressBar1.Maximum Then
                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                End If
            End If
        Next

        Me.ProgressBar1.Value = Me.ProgressBar1.Maximum
    End Sub
    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Private Sub SelectXLSLookupFileToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SelectXLSLookupFileToolStripMenuItem.Click
        '##################################################################################################################################################
        '#      Function:       OpenPACPDatabaseToolStripMenuItem_Click()
        '#      Description:    Opens the file browse dialogue
        '#      Version:        1.0
        '#      Last changed:   22.06.2017
        '#      Author:         s.bregy@wincan.com
        '##################################################################################################################################################
        Me.OpenFileDialog1.Title = "Select XLS Lookup file"
        Me.OpenFileDialog1.Filter = "All files (*.XLS)|*.XLS|All files (*.XLS)|*.XLS"
        Me.ProgressBar1.Value = 0
        Me.OpenFileDialog1.RestoreDirectory = True

        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Panel1.Visible = True 'Panel1 shows the WinCan Logo and a label with a user information like "DB check in progress, please be patient"
            Me.Panel1.Refresh()
            Me.lblXLSLookUp.Text = Me.OpenFileDialog1.FileName

            Me.Panel1.Visible = True
            Me.lblStatus.Text = "Loading data. Please be patient ..."

            GetXLSData(Me.lblXLSLookUp.Text)
            Me.Panel1.Visible = False 'Hide Panel1 again
            If Me.lblFilePath.Text <> "-" Then
                Me.cmdUpdate.Enabled = True
                Me.cmdCopyMedia.Enabled = True
            End If
        End If
        Me.ProgressBar1.Value = 0
    End Sub

    Private Function GetAssetNumber(sID As String) As String

        For Each item As KeyValuePair(Of String, String) In XLSDictionary
            If item.Value = sID Then
                GetAssetNumber = item.Key
            End If
        Next

        If GetAssetNumber = "" Then
            GetAssetNumber = "N/A"
        End If

    End Function

    Private Sub UpdatePACPData(sAsset As String, sID As String)
        Dim sSQL As String
        Dim sConnection As String
        sSQL = "UPDATE Inspections SET Inspections.Pipe_Segment_Reference='" & sAsset & "' WHERE Inspections.InspectionID=" & sID
        sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"
        Using connection As New OleDbConnection(sConnection)
            Dim command As New OleDbCommand(sSQL)
            Dim dAdapter As OleDbDataAdapter
            dAdapter = New OleDbDataAdapter()
            dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

            command.Connection = connection

            connection.Open()
            command.ExecuteNonQuery()

            dAdapter.Dispose()
            command.Dispose()
            connection.Close() 'Closes the connection
        End Using
    End Sub

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs)
        Dim sAssetNumber As String

        Me.cmdUpdate.Enabled = False
        Me.cmdCopyMedia.Enabled = False
        Me.MenuStrip1.Enabled = False

        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Maximum = DataGridView1.Rows.Count

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.IsNewRow Then 'When reaching the empty row exit the loop (empty row is the very last one)
                Exit For
            End If
            row.Selected = True
            If Not row.Displayed = True Then
                Me.DataGridView1.FirstDisplayedScrollingRowIndex = row.Index 'if the row is outside the visible area, make it visible
            End If
            row.Selected = False
            row.Cells(7).Style.BackColor = Color.Yellow
            row.Cells(7).Value = "Searching ..."
            sAssetNumber = GetAssetNumber(row.Cells(1).Value.ToString)
            If sAssetNumber <> "N/A" Then
                UpdatePACPData(sAssetNumber, row.Cells(0).Value.ToString)
                row.Cells(7).Style.BackColor = Color.Green
                row.Cells(7).Value = "Updated (" & sAssetNumber & ")"
            Else
                row.Cells(7).Style.BackColor = Color.Red
                row.Cells(7).Value = "N/A"
            End If
            Application.DoEvents()
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next

        MsgBox("PACP database updated. Please check the log file", vbInformation, "PACP check")
        Me.ProgressBar1.Value = 0

        Me.cmdUpdate.Enabled = True
        Me.cmdCopyMedia.Enabled = True
        Me.MenuStrip1.Enabled = True
    End Sub

    Private Function CopyMediaFile(sSource As String, sTarget As String) As String
        CopyMediaFile = ""

        Try
            If File.Exists(sSource) = False Then
                CopyMediaFile = "Source file does not exists"
                Exit Function
            End If
            If File.Exists(sTarget) = True Then
                CopyMediaFile = "Target file already exists"
                Exit Function
            End If
            File.Copy(sSource, sTarget)
            CopyMediaFile = "File copied"
        Catch ex As Exception
            CopyMediaFile = "Copy failed"
        End Try
    End Function

    Private Sub cmdCopyMedia_Click(sender As System.Object, e As System.EventArgs) Handles cmdCopyMedia.Click
        Me.cmdUpdate.Enabled = False
        Me.cmdCopyMedia.Enabled = False
        Me.MenuStrip1.Enabled = False
        Dim sSrcFile As String
        Dim sTarFile As String
        Dim sStatus As String

        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Maximum = DataGridView2.Rows.Count

        For Each row As DataGridViewRow In DataGridView2.Rows
            If row.IsNewRow Then 'When reaching the empty row exit the loop (empty row is the very last one)
                Exit For
            End If
            row.Selected = True
            If Not row.Displayed = True Then
                Me.DataGridView2.FirstDisplayedScrollingRowIndex = row.Index 'if the row is outside the visible area, make it visible
            End If
            row.Selected = False
            row.Cells(3).Style.BackColor = Color.Yellow
            row.Cells(3).Value = "Copy file ..."
            Application.DoEvents()

            sSrcFile = Me.lblMediaPath.Text & "\" & row.Cells(2).Value.ToString & row.Cells(1).Value.ToString
            sTarFile = Me.lblMediaTargetPath.Text & "\Picture\Sec\" & row.Cells(1).Value.ToString

            sStatus = CopyMediaFile(sSrcFile, sTarFile)
            If sStatus = "File copied" Then
                row.Cells(3).Style.BackColor = Color.Green
                row.Cells(3).Value = "File copied"
                'ResetMediaPath(row.Cells(0).Value.ToString, "Photo")
            ElseIf sStatus = "Target file already exists" Then
                row.Cells(3).Style.BackColor = Color.Orange
                row.Cells(3).Value = sStatus
                'ResetMediaPath(row.Cells(0).Value.ToString, "Photo")
            Else
                row.Cells(3).Style.BackColor = Color.Red
                row.Cells(3).Value = sStatus
            End If

            sStatus = ""
            Application.DoEvents()
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next


        'VIDEO

        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Maximum = DataGridView3.Rows.Count
        For Each row As DataGridViewRow In DataGridView3.Rows
            If row.IsNewRow Then 'When reaching the empty row exit the loop (empty row is the very last one)
                Exit For
            End If
            row.Selected = True
            If Not row.Displayed = True Then
                Me.DataGridView3.FirstDisplayedScrollingRowIndex = row.Index 'if the row is outside the visible area, make it visible
            End If
            row.Selected = False
            row.Cells(3).Style.BackColor = Color.Yellow
            row.Cells(3).Value = "Copy file ..."
            Application.DoEvents()

            If Strings.Left(row.Cells(2).Value.ToString, 1) <> "\" And Mid(row.Cells(2).Value.ToString, 2, 1) <> ":" Then
                sSrcFile = Me.lblVideoSourcePath.Text & "\" & row.Cells(2).Value.ToString & row.Cells(1).Value.ToString
                sTarFile = Me.lblMediaTargetPath.Text & "\Video\Sec\" & row.Cells(1).Value.ToString
            Else
                If Strings.Right(row.Cells(2).Value.ToString, 1) = "\" Then
                    sSrcFile = row.Cells(2).Value.ToString & row.Cells(1).Value.ToString
                    sTarFile = Me.lblMediaTargetPath.Text & "\Video\Sec\" & row.Cells(1).Value.ToString
                Else
                    sSrcFile = row.Cells(2).Value.ToString & "\" & row.Cells(1).Value.ToString
                    sTarFile = Me.lblMediaTargetPath.Text & "\Video\Sec\" & row.Cells(1).Value.ToString
                End If
            End If

            sStatus = CopyMediaFile(sSrcFile, sTarFile)
            If sStatus = "File copied" Then
                row.Cells(3).Style.BackColor = Color.Green
                row.Cells(3).Value = "File copied"
                'ResetMediaPath(row.Cells(0).Value.ToString, "Video", modInivb.sNewVidPath)
            ElseIf sStatus = "Target file already exists" Then
                row.Cells(3).Style.BackColor = Color.Orange
                row.Cells(3).Value = sStatus
                'ResetMediaPath(row.Cells(0).Value.ToString, "Video", modInivb.sNewVidPath)
            Else
                row.Cells(3).Style.BackColor = Color.Red
                row.Cells(3).Value = sStatus
            End If

            sStatus = ""
            Application.DoEvents()
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next

        '////VIDEO

        MsgBox("Media files copied. Please check the log file", vbInformation, "Media Copy")
        Me.ProgressBar1.Value = 0

        Me.cmdUpdate.Enabled = True
        Me.cmdCopyMedia.Enabled = True
        Me.MenuStrip1.Enabled = True
    End Sub

    Private Sub ResetMediaPath(sID As String, sType As String, sVal As String)
        Dim sSQL As String
        Dim sConnection As String

        Select Case sType
            Case "Photo"
                sSQL = "UPDATE Media_Conditions SET Image_Path = NULL WHERE MediaCondID=" & sID
                sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"
                Using connection As New OleDbConnection(sConnection)
                    Dim command As New OleDbCommand(sSQL)
                    Dim dAdapter As OleDbDataAdapter
                    dAdapter = New OleDbDataAdapter()
                    dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

                    command.Connection = connection

                    connection.Open()
                    command.ExecuteNonQuery()

                    dAdapter.Dispose()
                    command.Dispose()
                    connection.Close() 'Closes the connection
                End Using
            Case "Video"
                sSQL = "UPDATE Media_Inspections SET Video_Location = '" & sVal & "' WHERE MediaID=" & sID
                sConnection = "Provider=microsoft.Jet.oledb.4.0;Data Source=" & Me.lblFilePath.Text & ";Persist Security Info=False;"
                Using connection As New OleDbConnection(sConnection)
                    Dim command As New OleDbCommand(sSQL)
                    Dim dAdapter As OleDbDataAdapter
                    dAdapter = New OleDbDataAdapter()
                    dAdapter.SelectCommand = New OleDbCommand(sSQL, connection)

                    command.Connection = connection

                    connection.Open()
                    command.ExecuteNonQuery()

                    dAdapter.Dispose()
                    command.Dispose()
                    connection.Close() 'Closes the connection
                End Using
        End Select
    End Sub

    Private Sub cmdSelectMediaPath_Click(sender As System.Object, e As System.EventArgs) Handles cmdSelectMediaPath.Click
        If ReadIni(modInivb.sConfig, "Config", "InitPathPhotoSource", "") <> "" Then
            FolderBrowserDialog1.SelectedPath = ReadIni(modInivb.sConfig, "Config", "InitPathPhotoSource", "")
        Else
            FolderBrowserDialog1.SelectedPath = Path.GetDirectoryName(lblFilePath.Text)
        End If

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.lblMediaPath.Text = FolderBrowserDialog1.SelectedPath
            If Me.lblMediaTargetPath.Text <> "-" And Me.lblMediaPath.Text <> "-" Then
                Me.cmdCopyMedia.Enabled = True
            Else
                Me.cmdCopyMedia.Enabled = False
            End If
        End If
    End Sub

    Private Sub cmdUpdate_Click_1(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        Dim sAssetNumber As String

        Me.cmdUpdate.Enabled = False
        Me.MenuStrip1.Enabled = False
        Me.cmdCopyMedia.Enabled = False

        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Maximum = DataGridView1.Rows.Count

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.IsNewRow Then 'When reaching the empty row exit the loop (empty row is the very last one)
                Exit For
            End If
            row.Selected = True
            If Not row.Displayed = True Then
                Me.DataGridView1.FirstDisplayedScrollingRowIndex = row.Index 'if the row is outside the visible area, make it visible
            End If
            row.Selected = False
            row.Cells(7).Style.BackColor = Color.Yellow
            row.Cells(7).Value = "Searching ..."
            sAssetNumber = GetAssetNumber(row.Cells(1).Value.ToString)
            If sAssetNumber <> "N/A" Then
                UpdatePACPData(sAssetNumber, row.Cells(0).Value.ToString)
                row.Cells(7).Style.BackColor = Color.Green
                row.Cells(7).Value = "Updated (" & sAssetNumber & ")"
            Else
                row.Cells(7).Style.BackColor = Color.Red
                row.Cells(7).Value = "N/A"
            End If
            Application.DoEvents()
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        Next

        MsgBox("PACP database updated. Please check the log file", vbInformation, "PACP check")
        Me.ProgressBar1.Value = 0

        Me.cmdUpdate.Enabled = True
        Me.MenuStrip1.Enabled = True
        Me.cmdCopyMedia.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If ReadIni(modInivb.sConfig, "Config", "InitPathMediaTarget", "") <> "" Then
            FolderBrowserDialog1.SelectedPath = ReadIni(modInivb.sConfig, "Config", "InitPathMediaTarget", "")
        End If
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.lblMediaTargetPath.Text = FolderBrowserDialog1.SelectedPath
            writeIni(modInivb.sConfig, "Config", "MediaTarget", Me.lblMediaTargetPath.Text)
            If Me.lblMediaTargetPath.Text <> "-" And Me.lblMediaPath.Text <> "-" Then
                Me.cmdCopyMedia.Enabled = True
            Else
                Me.cmdCopyMedia.Enabled = False
            End If
        End If
    End Sub

    Private Sub FileToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FileToolStripMenuItem.Click

    End Sub

    Private Sub frmMain_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        modInivb.SetConfigPath()
        Me.lblMediaTargetPath.Text = ReadIni(modInivb.sConfig, "Config", "MediaTarget", "")
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick1(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            If row.IsNewRow Then Exit For
            If row.Cells(7).Value.ToString = "N/A" Then
                row.Cells(7).Style.BackColor = Color.Red
            Else
                row.Cells(7).Style.BackColor = Color.Green
            End If
        Next
    End Sub

    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        For Each row As DataGridViewRow In Me.DataGridView2.Rows
            If row.IsNewRow Then Exit For
            If row.Cells(3).Value.ToString = "File copied" Then
                row.Cells(3).Style.BackColor = Color.Green
            ElseIf row.Cells(3).Value.ToString = "Target file already exists" Then
                row.Cells(3).Style.BackColor = Color.Orange
            Else
                row.Cells(3).Style.BackColor = Color.Red
            End If
        Next
    End Sub

    Private Sub DataGridView3_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.ColumnHeaderMouseClick
        For Each row As DataGridViewRow In Me.DataGridView3.Rows
            If row.IsNewRow Then Exit For
            If row.Cells(3).Value.ToString = "File copied" Then
                row.Cells(3).Style.BackColor = Color.Green
            ElseIf row.Cells(3).Value.ToString = "Target file already exists" Then
                row.Cells(3).Style.BackColor = Color.Orange
            Else
                row.Cells(3).Style.BackColor = Color.Red
            End If
        Next
    End Sub

    Private Sub cmdSelectVideoPath_Click(sender As System.Object, e As System.EventArgs) Handles cmdSelectVideoPath.Click
        If ReadIni(modInivb.sConfig, "Config", "InitPathVideoSource", "") <> "" Then
            FolderBrowserDialog1.SelectedPath = ReadIni(modInivb.sConfig, "Config", "InitPathVideoSource", "")
        End If

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.lblVideoSourcePath.Text = FolderBrowserDialog1.SelectedPath
            If Me.lblVideoSourcePath.Text <> "-" And Me.lblVideoSourcePath.Text <> "-" Then
                Me.cmdCopyMedia.Enabled = True
            Else
                Me.cmdCopyMedia.Enabled = False
            End If
        End If
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
