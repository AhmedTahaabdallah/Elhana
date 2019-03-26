Imports System.Data.SqlClient

Public Class Form1
    Dim sqlconn As New SqlConnection With {.ConnectionString = "Server=PC-PC\SQLEXPRESS;Database=master;User=sa;Pwd=141751527"}

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SaveFileDialog1.FileName = "Elhana " + DateAndTime.DateString
        SaveFileDialog1.Filter = "SQL Server database backup files|*.bak"
        'SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                Me.Cursor = Cursors.WaitCursor
                'Dim cmd As New SqlCommand("BACKUP DATABASE Elhana TO disk='" & SaveFileDialog1.FileName & "' WITH COPY_ONLY", sqlconn)
                Dim cmd As New SqlCommand("BACKUP DATABASE Elhana TO disk='" & SaveFileDialog1.FileName & "'", sqlconn)
                If sqlconn.State = 1 Then sqlconn.Close()
                sqlconn.Open()
                cmd.ExecuteNonQuery()
                sqlconn.Close()
                Me.Cursor = Cursors.Default
                MsgBox("تم عمل باك اب لقاعدة البيانات بنجاح", MsgBoxStyle.Information, "باك اب")
                'MessageBox.Show("back up saved")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        OpenFileDialog1.Filter = "SQL Server database backup files|*.bak"
        'OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                'Dim cmd As New SqlCommand("RESTORE DATABASE Elhana FROM  disk='" & OpenFileDialog1.FileName & "' WITH REPLACE", sqlconn)
                'MessageBox.Show(OpenFileDialog1.FileName)
                Me.Cursor = Cursors.WaitCursor
                Dim cmd As New SqlCommand("RESTORE DATABASE Elhana FROM  disk='" & OpenFileDialog1.FileName & "'", sqlconn)
                If sqlconn.State = 1 Then sqlconn.Close()
                sqlconn.Open()
                cmd.ExecuteNonQuery()
                sqlconn.Close()
                Me.Cursor = Cursors.Default
                MsgBox("تم استرجاع قاعدة البيانات بنجاح", MsgBoxStyle.Information, "استرجاع")
                'MessageBox.Show("RESTORED")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                MsgBox("قم باغلاق البرنامج الرئيسى ثم العودة هنا", MsgBoxStyle.Information, "خطأ")
                End
            End Try
        End If
    End Sub

   
End Class
