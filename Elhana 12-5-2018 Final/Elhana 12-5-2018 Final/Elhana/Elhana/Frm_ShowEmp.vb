Imports ADODB
Imports System.Data.OleDb
Public Class Frm_ShowEmp
    'Public gcon As New OleDbConnection
    Private Sub Frm_ShowEmp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=AHMED1"
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
            DataGridView1.Columns(1).Width = 150
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 140
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 140
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 220
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            If u_Id = 1 Then
                DataGridView1.Columns(8).Width = 140
                DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(9).Width = 160
                DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
                stt = "SELECT em.Emp_Name as [أسم الموظف], em.Emp_Phone as [رقم الموبايل], em.Emp_Datework as [تاريخ بداية العمل], em.Emp_Salary as [الراتب], em.adress as [العنوان], em.Emp_age as [السن], em.Emp_State as [يعمل حاليا بالمضرب], em.Emp_Salry_State as [على قوة الراتب], em.Emp_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=em.User_ID) as [مستخدم الأضافة], (Select usw.User_Name From Users as usw Where usw.User_ID=em.User_ID_edit) as [مستخدم التعديل] FROM Employyes as em where em.Emp_Flag='" & s & "' ORDER BY em.Emp_Id ASC"
            Else
                stt = "SELECT Emp_Name as [أسم الموظف], Emp_Phone as [رقم الموبايل], Emp_Datework as [تاريخ بداية العمل], Emp_Salary as [الراتب], adress as [العنوان], Emp_age as [السن], Emp_State as [يعمل حاليا بالمضرب], Emp_Salry_State as [على قوة الراتب], Emp_Note as [ملاحظات] FROM Employyes where Emp_Flag='" & s & "' ORDER BY Emp_Id ASC"
            End If


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