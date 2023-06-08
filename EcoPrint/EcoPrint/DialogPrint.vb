Imports System.Windows.Forms

Public Class DialogPrint

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Select Case Me.Text
            Case "Entering a task (New)*"

                If Len(TextBox1.Text) < 1 Then
                    MsgBox("Enter the number of pages", vbExclamation, "Error")
                Else
                    Dim row As String() = New String() {TextBox1.Text, ComboBox1.Text}
                    Form1.DataGridView1.Rows.Add(row)
                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Me.Close()
                End If



        End Select

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case Me.Text
            Case "Entering a task (New)*"
                TextBox1.Select()
                TextBox1.Text = ""
                ComboBox1.Text = "one- sided"
        End Select
    End Sub
End Class
