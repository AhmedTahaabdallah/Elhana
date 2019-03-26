Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class Frm_Store

    'Public gcon As New OleDbConnection

    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet

    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 160
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            stt = ""
            stt = "SELECT (SELECT it.Item_Name FROM Itemes as it where it.Item_ID=sq.Item_ID And it.Item_DelFlag='" & s & "') as [أسم الصنف], sq.Store_Quntity as [الكمية], sq.Item_Wight_Type as [نوع الصنف] FROM Store_Quentity as sq where sq.Store_DelFlag='" & s & "' ORDER BY sq.Store_ID ASC"

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            dgv()
            DataGridView1.Refresh()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Count(Store_ID) as c_No FROM Store_Quentity where Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreVend("c_No").Value
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Frm_Store_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If gcon.State = 1 Then gcon.Close()
            gcon.Open()
            get_all()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Cursor = Cursors.WaitCursor
        Dim rpt As New Rpt_Showstorquintiy
        Dim frm_rpt As New frm_rpt_showstorquintiy
        rdat.Clear()
        rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
        rcon.Open()
        Dim ssd As Boolean
        ssd = False
        
        rcomand = New SqlCommand("SELECT (SELECT it.Item_Name FROM Itemes as it where it.Item_ID=sq.Item_ID And it.Item_DelFlag='" & ssd & "') as itemname, sq.Store_Quntity as quntity, sq.Item_Wight_Type as ty FROM Store_Quentity as sq where sq.Store_DelFlag='" & ssd & "' ORDER BY sq.Store_ID ASC", rcon)
        
        rada = New SqlDataAdapter(rcomand)
        
        rada.Fill(rdat, "Store_Quentity")
        rpt.SetDataSource(rdat)
        
        frm_rpt.CrystalReportViewer1.ReportSource = rpt
        frm_rpt.CrystalReportViewer1.Refresh()
        frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
        Dim frm As New Form
        With frm
            .Controls.Add(frm_rpt.CrystalReportViewer1)
            .Text = "طباعة المخزن"
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With

        Me.Cursor = Cursors.Default
    End Sub
End Class