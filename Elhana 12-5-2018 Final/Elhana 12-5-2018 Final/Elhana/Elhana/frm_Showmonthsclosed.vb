Imports System.Data.OleDb
Public Class frm_Showmonthsclosed

    Private Sub frm_Showmonthsclosed_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            DataGridView1.Columns(0).Width = 110
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 110
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 110
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 110
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 110
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 130
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).Width = 150
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(10).Width = 100
            DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(11).Width = 110
            DataGridView1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(12).Width = 150
            DataGridView1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(13).Width = 150
            DataGridView1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(14).Width = 230
            DataGridView1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            stt = "Select (Select ym.Month_Name from year_monthes as ym where ym.Month_ID = mc.Month_ID) as [الشهر الحالى],mc.mc_totalAllquntityprice as [إجمالى البضاعة],mc.mc_totalblance as [إجمالى الخزنة], mc.mc_totalallcustmoney as [إجمالى الأجل], mc.mc_TotalBlasticPrice as [إجمالى سعر البلاستيك], mc.mc_totalEslam as [مسحوبات أسلام], mc.mc_totalAhmed as [مسحوبات أحمد], mc.mc_totalbeforefinal as [الإجمالى], mc.mc_totalvendblance as [إجمالى اللى علينا], mc.mc_firsttotal as [الإجمالى ], (Select ym1.Month_Name from year_monthes as ym1 where ym1.Month_ID = mc.Month_IDlast) as [الشهر السابق], mc.mc_lasttotalblance as [إجمالى الرصيد], mc.mc_totalbenfet as [إجمالى ربح الشهر الحالى], mc.mc_Finaltotalblance as [إجمالى رصيد الشهر الحالى], mc.mc_notes as [ملاحظات] FROM month_closed as mc where mc.mc_DeleteFlag='" & s & "' ORDER BY mc.Month_ID ASC"
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