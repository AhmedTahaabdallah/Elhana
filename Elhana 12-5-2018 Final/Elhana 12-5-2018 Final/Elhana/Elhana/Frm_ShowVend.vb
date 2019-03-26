Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class Frm_ShowVend

    'Public gcon As New OleDbConnection
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet
    Private Sub Frm_ShowVend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\AHMED1"
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
            DataGridView1.Columns(0).Width = 180
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 110
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 110
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 110
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 190
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 230
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            If u_Id = 1 Then
                DataGridView1.Columns(6).Width = 150
                DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(7).Width = 160
                DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

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

            If u_Id = 1 Then
                stt = "SELECT vnd.Vend_Name as [أسم المورد],vnd.Vend_Blance_Type as [نوع الرصيد], vnd.Vend_Blance_Dein as [الرصيد], vnd.Vend_Phone as [رقم الموبايل], vnd.Vend_Address as [العنوان], vnd.Vend_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=vnd.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=vnd.User_ID_edit) as [مستخدم التعديل] FROM Vendores as vnd where vnd.Vend_DelFlag='" & s & "' And vnd.Vend_visable='" & s & "' ORDER BY vnd.Vend_ID ASC"
            Else
                stt = "SELECT Vend_Name as [أسم المورد],Vend_Blance_Type as [نوع الرصيد], Vend_Blance_Dein as [الرصيد], Vend_Phone as [رقم الموبايل], Vend_Address as [العنوان], Vend_Note as [ملاحظات] FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' ORDER BY Vend_ID ASC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            dgv()
            DataGridView1.Refresh()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Count(Vend_ID) as c_No FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreVend("c_No").Value
            End If
            Dim li As Integer
            Dim al As Integer
            Dim fg As String
            fg = "defulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Vend_Blance_Dein) as c_No FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' And Vend_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label4.Text = 0
                al = 0
            Else
                Label4.Text = rs_StoreVend("c_No").Value
                al = rs_StoreVend("c_No").Value
            End If
            fg = "notdefulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Vend_Blance_Dein) as c_No FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' And Vend_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label7.Text = 0
                li = 0
            Else
                Label7.Text = rs_StoreVend("c_No").Value
                li = rs_StoreVend("c_No").Value
            End If
            Label10.Text = al - li
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Cursor = Cursors.WaitCursor
        Dim rpt As New Rpt_Showvendors
        Dim frm_rpt As New frm_rpt_showvendors
        rdat.Clear()
        rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
        rcon.Open()
        Dim ssd As Boolean
        ssd = False

        rcomand = New SqlCommand("SELECT vnd.Vend_Name as vendname,vnd.Vend_Blance_Type vendtype, vnd.Vend_Blance_Dein as vendblance, vnd.Vend_Note as vendnote FROM Vendores as vnd where vnd.Vend_DelFlag='" & ssd & "' And vnd.Vend_visable='" & ssd & "' ORDER BY vnd.Vend_ID ASC", rcon)

        rada = New SqlDataAdapter(rcomand)

        rada.Fill(rdat, "vendors")
        rpt.SetDataSource(rdat)

        frm_rpt.CrystalReportViewer1.ReportSource = rpt
        frm_rpt.CrystalReportViewer1.Refresh()
        frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
        Dim frm As New Form
        With frm
            .Controls.Add(frm_rpt.CrystalReportViewer1)
            .Text = "طباعة كل الموردين"
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With

        Me.Cursor = Cursors.Default
    End Sub
End Class