Imports System.Data.OleDb
Public Class Frm_ShowItems

    'Public gcon As New OleDbConnection

    Private Sub Frm_ShowVend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\AHMED1"
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
            DataGridView1.Columns(0).Width = 170
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 250
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            If u_Id = 1 Then
                DataGridView1.Columns(4).Width = 150
                DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
                stt = "SELECT itm.Item_Name as [أسم الصنف], itm.Item_Price as [سعر الصنف], itm.Item_Wight_Type as [نوع الصنف], itm.Item_Notes as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=itm.User_ID) as [مستخدم الأضافة] FROM Itemes as itm where itm.Item_DelFlag='" & s & "' ORDER BY itm.Item_ID ASC"
            Else
                stt = "SELECT Item_Name as [أسم الصنف], Item_Price as [سعر الصنف], Item_Wight_Type as [نوع الصنف], Item_Notes as [ملاحظات] FROM Itemes where Item_DelFlag='" & s & "' ORDER BY Item_ID ASC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            dgv()
            DataGridView1.Refresh()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Count(Item_ID) as c_No FROM Itemes where Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreVend("c_No").Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub
End Class