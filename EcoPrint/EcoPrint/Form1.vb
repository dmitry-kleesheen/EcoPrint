Imports System.Drawing.Printing
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports System.Management
Imports System.IO

Public Class Form1
    Private Sub ButtonClose_Click(sender As Object, e As EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub Control_Disabled(ByVal ControlDisabled As Boolean)

        For Each ctrl In Me.Controls

            Select Case ControlDisabled
                Case True
                    If ctrl.Tag = "f" Then
                        ctrl.Enabled = False
                    End If
                Case False
                    If ctrl.Tag = "f" Then
                        ctrl.Enabled = True
                    End If
            End Select
        Next

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName) & " - Version: " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor

        Dim prninst As String
        Dim printdocument As New PrintDocument

        Dim oPS As New System.Drawing.Printing.PrinterSettings

        For i = 0 To PrinterSettings.InstalledPrinters.Count - 1
            prninst = PrinterSettings.InstalledPrinters(i).ToString
            CmbPrinter.Items.Add(prninst)

            If oPS.PrinterName = prninst Then
                CmbPrinter.SelectedIndex = i
            End If

        Next


    End Sub

    Private Sub ButtonOdnoadd_Click(sender As Object, e As EventArgs) Handles ButtonOdnoadd.Click

        DialogPrint.Text = "Entering a task (New)*"
        DialogPrint.ShowDialog()

    End Sub

    Private Sub ButtonOdnodelete_Click(sender As Object, e As EventArgs) Handles ButtonOdnodelete.Click

        Try
            If DataGridView1.CurrentRow.Index >= 0 Then

                DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
            End If
        Catch ex As Exception
            MsgBox("Select value to delete!", vbExclamation, "Error")
        End Try

    End Sub

    Private Sub ButtonPrint_Click(sender As Object, e As EventArgs) Handles ButtonPrint.Click

        If IO.File.Exists(TextBoxFile.Text) = True Then
            Print_Document()
        Else
            MsgBox("Select a file!", vbExclamation, "Error")
        End If

    End Sub

    Private Sub LinkLabelFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabelFile.LinkClicked
        'Dim myStream As Stream = Nothing

        Dim openFileDialog1 As New OpenFileDialog()


        openFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments


        openFileDialog1.Filter = "Files MS Word (*.doc,.docx, rtf)|*.docx;*.doc;*.rtf"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        fl_name = openFileDialog1.FileName.ToString

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try

                TextBoxFile.Text = ""
                TextBoxFile.Text = openFileDialog1.FileName.ToString

                DataGridView1.Rows.Clear()
                ListBoxLog.Items.Clear()

            Catch Ex As Exception
                MessageBox.Show("Error reading the file. Error Description: " & Ex.Message)
            Finally

            End Try
        End If
    End Sub

    Dim fl_name As String

    Private Sub Print_Document()

        ListBoxLog.Items.Clear()

        Control_Disabled(True)

        Cursor = Cursors.WaitCursor

        For i = 0 To DataGridView1.Rows.Count - 1

            Select Case DataGridView1.Item(1, i).Value

                Case "one- sided"

                    Dim missing As Object = System.Reflection.Missing.Value
                    Dim poApp As New Word.Application
                    Dim poDoc As Word.Document

                    GavPrinterSetting(CmbPrinter.SelectedIndex, 1)

                    poApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsMessageBox
                    poDoc = poApp.Documents.Open(TextBoxFile.Text)
                    poDoc.PrintOut(False, False, WdPrintOutRange.wdPrintRangeOfPages, missing, missing, missing, WdPrintOutItem.wdPrintDocumentContent, "1", DataGridView1.Item(0, i).Value.ToString, WdPrintOutPages.wdPrintAllPages, missing, False, missing, False, 0, 0, 0, 0)
                    poDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                    'clean up
                    poDoc = Nothing
                    'close word
                    poApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges)
                    poApp = Nothing
                    ListBoxLog.Items.Add("One-sided printing - " & DataGridView1.Item(0, i).Value.ToString)

                Case "duplex"

                    Dim missing As Object = System.Reflection.Missing.Value
                    Dim poApp As New Word.Application
                    Dim poDoc As Word.Document

                    GavPrinterSetting(CmbPrinter.SelectedIndex, 2)

                    poApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsMessageBox

                    poDoc = poApp.Documents.Open(TextBoxFile.Text)
                    poDoc.PrintOut(False, False, WdPrintOutRange.wdPrintRangeOfPages, missing, missing, missing, WdPrintOutItem.wdPrintDocumentContent, 1, DataGridView1.Item(0, i).Value.ToString, WdPrintOutPages.wdPrintAllPages, False, False, missing, False, 0, 0, 0, 0)
                    poDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                    'clean up
                    poDoc = Nothing
                    'close word
                    poApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges)
                    poApp = Nothing
                    ListBoxLog.Items.Add("Duplex printing - " & DataGridView1.Item(0, i).Value.ToString)

                    GavPrinterSetting(CmbPrinter.SelectedIndex, 1)

            End Select
        Next

        Cursor = Cursors.Default

        Control_Disabled(False)

        ListBoxLog.Items.Add("Document printing: (" & IO.Path.GetFileName(TextBoxFile.Text) & ") - is completed!")

    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub CmbPrinter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbPrinter.SelectedIndexChanged
        My.Settings.PRNSET = CmbPrinter.SelectedIndex
        My.Settings.Save()
    End Sub

    Private Sub TextBoxFile_DragEnter(sender As Object, e As DragEventArgs) Handles TextBoxFile.DragEnter
        If e.Data.GetDataPresent("*.*") Or e.Data.GetDataPresent(DataFormats.FileDrop) Then

            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim FormatFile As String

            FormatFile = Path.GetExtension(files(0).ToString).ToString

            If IO.File.Exists(files(0).ToString) = True Then

                Select Case FormatFile
                    Case ".doc"
                        e.Effect = DragDropEffects.Copy
                        TextBoxFile.Text = files(0).ToString
                    Case ".docx"
                        e.Effect = DragDropEffects.Copy
                        TextBoxFile.Text = files(0).ToString
                    Case ".rtf"
                        e.Effect = DragDropEffects.Copy
                        TextBoxFile.Text = files(0).ToString
                    Case Else
                        e.Effect = DragDropEffects.None
                End Select

            Else
                e.Effect = DragDropEffects.None
                MsgBox("File is not found!" & vbCrLf & files(0).ToString, MsgBoxStyle.Critical, "Error")
            End If

        End If
    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting

        Select Case DataGridView1.Rows.Count
            Case 1
                ButtonOdnodelete.Enabled = True
                ButtonPrint.Enabled = True
                ButtonDOWN.Enabled = False
                ButtonUP.Enabled = False
            Case 0
                ButtonPrint.Enabled = False
                ButtonOdnodelete.Enabled = False
                ButtonDOWN.Enabled = False
                ButtonUP.Enabled = False
            Case >= 2
                ButtonDOWN.Enabled = True
                ButtonUP.Enabled = True
                ButtonOdnodelete.Enabled = True
                ButtonOdnodelete.Enabled = True
        End Select

    End Sub

    Dim rowIndex As Integer

    Private Sub ButtonUP_Click(sender As Object, e As EventArgs) Handles ButtonUP.Click
        ' get the index of selected row
        rowIndex = DataGridView1.SelectedCells(0).OwningRow.Index

        Dim item As String() = {DataGridView1.Item(0, rowIndex).Value, DataGridView1.Item(1, rowIndex).Value}

        If rowIndex = DataGridView1.Rows.Count - 1 Or rowIndex >= 1 Then
            DataGridView1.Rows.RemoveAt(rowIndex)
            DataGridView1.Rows.Insert(rowIndex - 1, item)
            DataGridView1.Rows(rowIndex - 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(rowIndex - 1).Cells(0)
        End If

    End Sub

    Private Sub ButtonDOWN_Click(sender As Object, e As EventArgs) Handles ButtonDOWN.Click
        rowIndex = DataGridView1.SelectedCells(0).OwningRow.Index

        Dim item As String() = {DataGridView1.Item(0, rowIndex).Value, DataGridView1.Item(1, rowIndex).Value}


        If rowIndex < DataGridView1.Rows.Count - 1 Then
            DataGridView1.Rows.RemoveAt(rowIndex)
            DataGridView1.Rows.Insert(rowIndex + 1, item)
            DataGridView1.Rows(rowIndex + 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(rowIndex + 1).Cells(0)
        End If

    End Sub

End Class
