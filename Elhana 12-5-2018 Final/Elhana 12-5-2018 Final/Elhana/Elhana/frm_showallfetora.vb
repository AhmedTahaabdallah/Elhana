Imports ADODB
Imports System.Data.OleDb
Public Class frm_showallfetora

    Private Sub frm_showallfetora_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If gcon.State = 1 Then gcon.Close()
            gcon.Open()
            get_all()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 70
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 105
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 100
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub
    Private Sub get_all()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String


            stt = "SELECT fo.Fetora_id as [رقم الفاتورة], cu.Cust_Name as [أسم العميل], fo.Fetora_date as [تاريخ الفاتورة ], fo.Fetora_blancebeforetype as [نوع قبل], fo.Fetora_blancebefore as [رصيد قبل], fo.Fetora_totalprice as [إجمالى الفاتورة], fo.Fetora_totalpay as [إجمالى المدفوع], fo.Fetora_blanceaftertype as [نوع بعد], fo.Fetora_blanceafter as [رصيد بعد] FROM Fetora as fo,Customers as cu where cu.Cust_DelFlag='" & s & "' And cu.Cust_ID=fo.Fetora_custid ORDER BY fo.Fetora_id DESC"



            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            dgv()
            DataGridView1.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class