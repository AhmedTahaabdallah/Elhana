Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class Frm_ShowStoreVend
    Public ite_id33 As Integer
    Public ite_id As Integer
    Public vend_id As Integer
    Public cust_id As Integer
    Public vend_id33 As Integer
    Public vend_id22 As Integer
    Public cust_id22 As Integer
    Public cust_id33 As Integer
    Public vend_id44 As Integer
    Public cust_id44 As Integer
    Public car_id As Integer
    Public gov_Name As String
    Public Dri_Name As String
    Public d1 As Date
    Public d2 As Date
    Public d3 As Date
    'Public gcon As New OleDbConnection
    Public month_id As Integer
    Public month_id33 As Integer
    Public month_id22 As Integer
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet
    Public sql_str As String
    Public mo As String
    Public pr As String
    Public V_Name As String
    Public T_Name As String
    Public C_Name As String
    Public V_fato As Integer
    Public T_fato As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    Public dstri2 As String

    Private Sub Frm_ShowStoreVend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\AHMED1"
        Try
            auth_report = rs_auth("User_storvend_report").Value
            If gcon.State = 1 Then gcon.Close()
            gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"
            gcon.Open()
            'get_all()
            'pr = "All"
            'T_Name = TextBox2.Text
            'Button2.Visible = True
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox5.Visible = False
            ComboBox6.Visible = False
            ComboBox7.Visible = False
            DateTimePicker1.Visible = True
            DateTimePicker3.Visible = False
            DateTimePicker2.Visible = False
            ComboBox4.Visible = False
            Button2.Visible = False
            ComboBox8.Visible = False
            ComboBox9.Visible = False
            ComboBox1.Text = "تاريخ الأذن"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub


    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 190
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 160
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 120
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 120
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).Width = 330
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            If u_Id = 1 Then
                DataGridView1.Columns(10).Width = 150
                DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(11).Width = 160
                DataGridView1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub dgvall()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 50
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 120
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 120
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 70
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 70
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).Width = 80
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(11).Width = 200
            DataGridView1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            If u_Id = 1 Then
                DataGridView1.Columns(12).Width = 100
                DataGridView1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(13).Width = 100
                DataGridView1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub dgvfatora()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 190
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 160
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 120
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 120
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 330
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            If u_Id = 1 Then
                DataGridView1.Columns(9).Width = 150
                DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(10).Width = 160
                DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String

            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If


            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvall()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_allitem()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvall()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_allitem_month()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID =" & month_id & " And st.Month_ID =" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID =" & month_id & " And st.Month_ID =" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvall()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID =" & month_id & " And st.Month_ID =" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID =" & month_id & " And st.Month_ID =" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub getItem_vendname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub getItem_custname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = " & ite_id & " And itm.Item_ID = " & ite_id & " And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_allVend()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            'Dim s7 As Boolean
            's7 = True
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            Dim rty As Integer
            rty = 1
            'stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
            
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_allCust()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            'Dim s7 As Boolean
            's7 = True
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            Dim rty As Integer
            rty = 1
            'stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
           
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_fatoraVend()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            dy = DateTimePicker3.Value.Year
            dm = DateTimePicker3.Value.Month
            dd = DateTimePicker3.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvfatora()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT SUM(st.Stor_Total_Price) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    T_fato = 0
            'Else
            '    T_fato = rs_StoreVend("c_No").Value
            'End If

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT SUM(st.Stor_Payment) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    V_fato = 0
            'Else
            '    V_fato = rs_StoreVend("c_No").Value
            'End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_fatoraCust()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            dy = DateTimePicker3.Value.Year
            dm = DateTimePicker3.Value.Month
            dd = DateTimePicker3.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvfatora()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT SUM(st.Stor_Total_Price) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    T_fato = 0
            'Else
            '    T_fato = rs_StoreVend("c_No").Value
            'End If

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT SUM(st.Stor_Payment) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    V_fato = 0
            'Else
            '    V_fato = rs_StoreVend("c_No").Value
            'End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_Datestri='" & dstri & "' And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub getvendname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub getcustname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_month_cust()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            'stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], cr.Car_No as [رقم السيارة], st.Drivers_Name as [اسم السائق], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_From as [من], st.Stor_To as [الى], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.car_Id=cr.car_Id And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Stor_DelFlag='" & s & "' And mn.Month_DeleFlag='" & s & " And st.Stor_Type='" & order_Type & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
           
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            rs_StoreVend.Close()
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_month_vend()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            'stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], cr.Car_No as [رقم السيارة], st.Drivers_Name as [اسم السائق], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_From as [من], st.Stor_To as [الى], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.car_Id=cr.car_Id And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Stor_DelFlag='" & s & "' And mn.Month_DeleFlag='" & s & " And st.Stor_Type='" & order_Type & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
           
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            rs_StoreVend.Close()
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    'Private Sub getcarno()
    '    Dim ds As New DataSet
    '    Dim dt As New DataTable
    '    ds.Tables.Add(dt)
    '    Dim da As New OleDbDataAdapter
    '    Dim s As Boolean
    '    s = False
    '    Dim order_Type As String
    '    order_Type = "Vend"
    '    Dim stt As String
    '    stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], cr.Car_No as [رقم السيارة], st.Drivers_Name as [اسم السائق], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_From as [من], st.Stor_To as [الى], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.car_Id=" & car_id & " And cr.car_Id=" & car_id & " And st.Vend_ID=vd.Vend_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
    '    da = New OleDbDataAdapter(stt, gcon)

    '    da.Fill(dt)
    '    DataGridView1.DataSource = dt.DefaultView
    '    gcon.Close()
    '    DataGridView1.Refresh()
    '    dgv()
    '    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
    '    rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.car_Id=" & car_id & " And cr.car_Id=" & car_id & " And st.Vend_ID=vd.Vend_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If IsDBNull(rs_StoreVend("c_No").Value) Then
    '        TextBox2.Text = 0
    '    Else
    '        TextBox2.Text = rs_StoreVend("c_No").Value
    '    End If
    '    'TextBox2.Text = rs_StoreVend("c_No").Value
    '    rs_StoreVend.Close()

    '    If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
    '    rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.car_Id=" & car_id & " And cr.car_Id=" & car_id & " And st.Vend_ID=vd.Vend_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If IsDBNull(rs_StoreCust("c_No").Value) Then
    '        Label3.Text = 0
    '    Else
    '        Label3.Text = rs_StoreCust("c_No").Value
    '    End If
    '    'Label3.Text = rs_StoreCust("c_No").Value
    '    rs_StoreCust.Close()
    'End Sub

    Private Sub get_datee()
        Try
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri2 = dy + "-" + dm + "-" + dd
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            'stt = "SELECT * FROM Stor where Stor_Date = #" & DateTimePicker1.Value.ToShortDateString() & "# And Stor_DelFlag =" & s & " And Stor_Type='" & order_Type & "'"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st , Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Datestri='" & dstri2 & "' And st.Vend_ID=vd.Vend_ID And st.Cust_ID=cu.Cust_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st , Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Datestri='" & dstri2 & "' And st.Vend_ID=vd.Vend_ID And st.Cust_ID=cu.Cust_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvall()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Stor_Datestri='" & dstri2 & "' And st.Vend_ID=vd.Vend_ID And st.Cust_ID=cu.Cust_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Stor_Datestri='" & dstri2 & "' And st.Vend_ID=vd.Vend_ID And st.Cust_ID=cu.Cust_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub


    'Private Sub get_datee()
    '    Dim ds As New DataSet
    '    Dim dt As New DataTable
    '    ds.Tables.Add(dt)
    '    Dim da As New OleDbDataAdapter
    '    Dim s As Boolean
    '    s = False
    '    Dim order_Type As String
    '    order_Type = "Vend"
    '    Dim stt As String
    '    'stt = "SELECT * FROM Stor where Stor_Date = #" & DateTimePicker1.Value.Date & "# And Stor_DelFlag =" & s & " And Stor_Type='" & order_Type & "'"
    '    stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], cr.Car_No as [رقم السيارة], st.Drivers_Name as [اسم السائق], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_From as [من], st.Stor_To as [الى], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Cars_Drivers as cr where st.Stor_Date = #" & DateTimePicker1.Value.Date & "# And st.Vend_ID = vd.Vend_ID And st.car_Id = cr.car_Id And st.Stor_DelFlag =" & s & " And st.Stor_Type='" & order_Type & "' And vd.Vend_DelFlag =" & s & " And cr.Car_DelFlag =" & s & " ORDER BY st.Stor_ID ASC"
    '    da = New OleDbDataAdapter(stt, gcon)

    '    da.Fill(dt)

    '    DataGridView1.DataSource = dt.DefaultView
    '    gcon.Close()
    '    DataGridView1.Refresh()
    '    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
    '    rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr where st.Stor_Date = #" & DateTimePicker1.Value.Date & "# And st.Vend_ID = vd.Vend_ID And st.car_Id = cr.car_Id And st.Stor_DelFlag =" & s & " And st.Stor_Type='" & order_Type & "' And vd.Vend_DelFlag =" & s & " And cr.Car_DelFlag =" & s & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If IsDBNull(rs_StoreVend("c_No").Value) Then
    '        TextBox2.Text = 0
    '    Else
    '        TextBox2.Text = rs_StoreVend("c_No").Value
    '    End If
    '    'TextBox2.Text = rs_StoreVend("c_No").Value
    '    rs_StoreVend.Close()

    '    If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
    '    rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr where st.Stor_Date = #" & DateTimePicker1.Value.Date & "# And st.Vend_ID = vd.Vend_ID And st.car_Id = cr.car_Id And st.Stor_DelFlag =" & s & " And st.Stor_Type='" & order_Type & "' And vd.Vend_DelFlag =" & s & " And cr.Car_DelFlag =" & s & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If IsDBNull(rs_StoreCust("c_No").Value) Then
    '        Label3.Text = 0
    '    Else
    '        Label3.Text = rs_StoreCust("c_No").Value
    '    End If
    '    'Label3.Text = rs_StoreCust("c_No").Value
    '    rs_StoreCust.Close()
    'End Sub

    'Private Sub get_drivername()
    '    Dim ds As New DataSet
    '    Dim dt As New DataTable
    '    ds.Tables.Add(dt)
    '    Dim da As New OleDbDataAdapter
    '    Dim s As Boolean
    '    s = False
    '    Dim order_Type As String
    '    order_Type = "Vend"
    '    Dim stt As String
    '    stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], cr.Car_No as [رقم السيارة], st.Drivers_Name as [اسم السائق], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_From as [من], st.Stor_To as [الى], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.Vend_ID=vd.Vend_ID And st.car_Id=cr.car_Id And st.Drivers_Name='" & Dri_Name & "' And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
    '    da = New OleDbDataAdapter(stt, gcon)

    '    da.Fill(dt)

    '    DataGridView1.DataSource = dt.DefaultView
    '    gcon.Close()
    '    DataGridView1.Refresh()
    '    dgv()
    '    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
    '    rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.car_Id = cr.car_Id And st.Drivers_Name='" & Dri_Name & "' And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '    If IsDBNull(rs_StoreVend("c_No").Value) Then
    '        TextBox2.Text = 0
    '    Else
    '        TextBox2.Text = rs_StoreVend("c_No").Value
    '    End If
    '    rs_StoreVend.Close()
    '    If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
    '    rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Cars_Drivers as cr, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.car_Id = cr.car_Id And st.Drivers_Name='" & Dri_Name & "' And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cr.Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If IsDBNull(rs_StoreCust("c_No").Value) Then
    '        Label3.Text = 0
    '    Else
    '        Label3.Text = rs_StoreCust("c_No").Value
    '    End If

    '    rs_StoreCust.Close()
    'End Sub

    Private Sub get_bettwntwodate()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            'stt = "SELECT * FROM Stor where Stor_Date Between  #" & DateTimePicker3.Value.Date & "# And #" & DateTimePicker2.Value.Date & "#"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Date BETWEEN " & DateTimePicker2.Value.ToShortDateString() & " And " & DateTimePicker3.Value.ToShortDateString() & " And st.Vend_ID=vd.Vend_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Date BETWEEN " & DateTimePicker2.Value.ToShortDateString() & " And " & DateTimePicker3.Value.ToShortDateString() & " And st.Vend_ID=vd.Vend_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Date BETWEEN " & DateTimePicker2.Value.ToShortDateString() & " And " & DateTimePicker3.Value.ToShortDateString() & " And st.Vend_ID=vd.Vend_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Stor_Date BETWEEN " & DateTimePicker2.Value.ToShortDateString() & " And " & DateTimePicker3.Value.ToShortDateString() & " And st.Vend_ID=vd.Vend_ID And st.Month_ID=mn.Month_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_monthname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgvall()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            rs_StoreVend.Close()
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Customers as cu, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_monthnamevend()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            'Dim s7 As Boolean
            's7 = True
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            Dim rty As Integer
            rty = 1
            'stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
           
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And vd.Vend_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_monthnamecust()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            'Dim s7 As Boolean
            's7 = True
            Dim z As Integer
            z = 0
            Dim order_Type As String
            order_Type = "Vend"
            Dim stt As String
            Dim rty As Integer
            rty = 1
            'stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Vend_Name as [أسم المورد], cu.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as cu, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID = vd.Vend_ID And st.Cust_ID = cu.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And cu.Cust_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            If u_Id = 1 Then
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات],(Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=st.User_ID_edit) as [مستخدم التعديل] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            Else
                stt = "SELECT st.Stor_ID as [رقم الأذن], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_Date DESC"
            End If
            
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd, Itemes as itm, year_monthes as mn where st.Cust_ID = vd.Cust_ID And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And st.Month_ID=" & month_id & " And mn.Month_ID=" & month_id & " And mn.Month_DeleFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' And vd.Cust_visable='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If
            'Label3.Text = rs_StoreCust("c_No").Value
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim uh As Boolean
            uh = True
            If ComboBox1.Text = "كل الاذونات" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_all()
                'T_Name = TextBox2.Text
                'pr = "All"
                'Button2.Visible = True
            End If
            If ComboBox1.Text = "كل أذونات الواردات للموردين" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_allVend()

                If auth_report = uh Then
                    T_Name = TextBox2.Text
                    pr = "AllVend"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    T_Name = TextBox2.Text
                    pr = "AllVend"
                    Button2.Visible = True
                End If

            End If
            If ComboBox1.Text = "كل أذونات الواردات للعملاء" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_allCust()
                If auth_report = uh Then
                    T_Name = TextBox2.Text
                    pr = "AllCust"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    T_Name = TextBox2.Text
                    pr = "AllCust"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "فاتورة باليوم لمورد" Then
                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_fatoraVend()
                If auth_report = uh Then
                    vend_id22 = vend_id
                    C_Name = ComboBox2.Text
                    pr = "fatora_Vend"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    vend_id22 = vend_id
                    C_Name = ComboBox2.Text
                    pr = "fatora_Vend"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "فاتورة باليوم لعميل" Then
                If ComboBox8.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox8.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_fatoraCust()
                If auth_report = uh Then
                    cust_id22 = cust_id
                    C_Name = ComboBox8.Text
                    pr = "fatora_Cust"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    cust_id22 = cust_id
                    C_Name = ComboBox8.Text
                    pr = "fatora_Cust"
                    Button2.Visible = True
                End If

            End If
            If ComboBox1.Text = "أسم المورد" Then
                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                getvendname()
                'T_Name = TextBox2.Text
                If auth_report = uh Then
                    vend_id33 = vend_id
                    V_Name = ComboBox2.Text
                    pr = "Vendor"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    vend_id33 = vend_id
                    V_Name = ComboBox2.Text
                    pr = "Vendor"
                    Button2.Visible = True
                End If
            End If

            If ComboBox1.Text = "أسم المورد والشهر" Then
                If ComboBox7.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox7.Select()
                    Exit Sub
                End If
                If ComboBox6.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox6.Select()
                    Exit Sub
                End If

                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_month_vend()
                'T_Name = TextBox2.Text
                If auth_report = uh Then
                    vend_id44 = vend_id
                    month_id22 = month_id
                    V_Name = ComboBox7.Text
                    mo = ComboBox6.Text
                    pr = "Vendor_month"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    vend_id44 = vend_id
                    month_id22 = month_id
                    V_Name = ComboBox7.Text
                    mo = ComboBox6.Text
                    pr = "Vendor_month"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "تاريخ الأذن" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_datee()
                'pr = "date"
                'Button2.Visible = True
            End If

            'If ComboBox1.Text = "بين تاريخين" Then

            '    If DateTimePicker2.Value.Date >= DateTimePicker3.Value.Date Then
            '        MsgBox("من فضلك أختر التاريخ الاقل", MsgBoxStyle.Information, "تنبيه")
            '        DateTimePicker2.Select()
            '        Exit Sub
            '    End If

            '    DataGridView1.DataSource = ""
            '    DataGridView1.Refresh()
            '    get_bettwntwodate()
            'End If
            If ComboBox1.Text = "الشهر" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_monthname()
                'T_Name = TextBox2.Text
                'month_id33 = month_id
                'mo = ComboBox4.Text
                'pr = "month"
                'Button2.Visible = True
            End If
            If ComboBox1.Text = "الشهر لواردات الموردين" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_monthnamevend()
                If auth_report = uh Then
                    T_Name = TextBox2.Text
                    month_id33 = month_id
                    mo = ComboBox4.Text
                    pr = "monthvend"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    T_Name = TextBox2.Text
                    month_id33 = month_id
                    mo = ComboBox4.Text
                    pr = "monthvend"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "الشهر لواردات العملاء" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_monthnamecust()
                If auth_report = uh Then
                    T_Name = TextBox2.Text
                    month_id33 = month_id
                    mo = ComboBox4.Text
                    pr = "monthcust"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    T_Name = TextBox2.Text
                    month_id33 = month_id
                    mo = ComboBox4.Text
                    pr = "monthcust"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم العميل" Then
                If ComboBox8.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox8.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                getcustname()
                If auth_report = uh Then
                    T_Name = TextBox2.Text
                    cust_id33 = cust_id
                    V_Name = ComboBox8.Text
                    pr = "Customer"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    T_Name = TextBox2.Text
                    cust_id33 = cust_id
                    V_Name = ComboBox8.Text
                    pr = "Customer"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم العميل والشهر" Then
                If ComboBox8.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox8.Select()
                    Exit Sub
                End If
                If ComboBox6.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox6.Select()
                    Exit Sub
                End If

                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_month_cust()
                'T_Name = TextBox2.Text
                If auth_report = uh Then
                    cust_id44 = cust_id
                    month_id22 = month_id
                    V_Name = ComboBox8.Text
                    mo = ComboBox6.Text
                    pr = "Customer_month"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    cust_id44 = cust_id
                    month_id22 = month_id
                    V_Name = ComboBox8.Text
                    mo = ComboBox6.Text
                    pr = "Customer_month"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم الصنف" Then
                If ComboBox9.Text = "" Then
                    MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox9.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_allitem()
                'T_Name = TextBox2.Text
                'ite_id33 = ite_id
                'V_Name = ComboBox9.Text
                'pr = "item"
                'Button2.Visible = True
            End If
            If ComboBox1.Text = "أسم الصنف والشهر" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                If ComboBox9.Text = "" Then
                    MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox9.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_allitem_month()
                'T_Name = TextBox2.Text
                'ite_id33 = ite_id
                'V_Name = ComboBox9.Text
                'pr = "item"
                'Button2.Visible = True
            End If
            If ComboBox1.Text = "أسم الصنف ومورد" Then
                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If
                If ComboBox9.Text = "" Then
                    MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox9.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                getItem_vendname()
                'T_Name = TextBox2.Text
                'vend_id33 = vend_id
                'V_Name = ComboBox2.Text
                'pr = "Vendor"
                'Button2.Visible = True
            End If
            If ComboBox1.Text = "أسم الصنف وعميل" Then
                If ComboBox8.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox8.Select()
                    Exit Sub
                End If
                If ComboBox9.Text = "" Then
                    MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox9.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                getItem_custname()
                'T_Name = TextBox2.Text
                'cust_id33 = cust_id
                'V_Name = ComboBox8.Text
                'pr = "Customer"
                'Button2.Visible = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Try
            Dim t As Integer
            t = 1000
            If Val(TextBox2.Text) > t Then
                Label15.Text = "طـن"
            ElseIf Val(TextBox2.Text) <= t Then
                Label15.Text = "كيلـو"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Vendores Where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox2.Items.Add(rs_Vendors("Vend_Name").Value)
                rs_Vendors.MoveNext()
            Loop
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox2.SelectedItem.ToString()
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            vend_id = rs_Vendors("Vend_ID").Value
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From Cars_Drivers Where Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox3.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox3.Items.Add(rs_Cars("Car_No").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox3.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From Cars_Drivers Where Car_No='" & u_name & "' And Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            car_id = rs_Cars("car_Id").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox5_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From Cars_Drivers Where Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox5.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox5.Items.Add(rs_Cars("Driver_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox5.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From Cars_Drivers Where Driver_Name='" & u_name & "' And Car_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            Dri_Name = rs_Cars("Driver_Name").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub




    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text = "كل الاذونات" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "كل أذونات الواردات للموردين" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "كل أذونات الواردات للعملاء" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "فاتورة باليوم لمورد" Then
                ComboBox2.Visible = True
                ComboBox2.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = True
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "فاتورة باليوم لعميل" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = True
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = True
                ComboBox8.Text = ""
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "أسم المورد" Then
                ComboBox8.Visible = False
                ComboBox2.Visible = True
                ComboBox2.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "أسم العميل" Then
                ComboBox8.Visible = True
                ComboBox8.Text = ""
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "أسم العميل والشهر" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = True
                ComboBox6.Text = ""
                ComboBox7.Visible = False
                ComboBox4.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = True
                ComboBox8.Text = ""
                ComboBox9.Visible = False
            End If
            'If ComboBox1.Text = "رقم السيارة" Then
            '    ComboBox2.Visible = False
            '    ComboBox3.Visible = True
            '    ComboBox6.Visible = False
            '    ComboBox7.Visible = False
            '    ComboBox5.Visible = False
            '    DateTimePicker1.Visible = False
            '    DateTimePicker3.Visible = False
            '    DateTimePicker2.Visible = False
            '    ComboBox4.Visible = False
            'End If
            If ComboBox1.Text = "أسم المورد والشهر" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = True
                ComboBox6.Text = ""
                ComboBox7.Visible = True
                ComboBox7.Text = ""
                ComboBox4.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "تاريخ الأذن" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ComboBox7.Visible = False
                DateTimePicker1.Visible = True
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox4.Visible = False
                Button2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            'If ComboBox1.Text = "أسم السائق" Then
            '    ComboBox2.Visible = False
            '    ComboBox3.Visible = False
            '    ComboBox5.Visible = True
            '    ComboBox6.Visible = False
            '    ComboBox7.Visible = False
            '    DateTimePicker1.Visible = False
            '    DateTimePicker3.Visible = False
            '    DateTimePicker2.Visible = False
            '    ComboBox4.Visible = False
            'End If
            'If ComboBox1.Text = "بين تاريخين" Then
            '    ComboBox2.Visible = False
            '    ComboBox6.Visible = False
            '    ComboBox7.Visible = False
            '    ComboBox3.Visible = False
            '    ComboBox5.Visible = False
            '    DateTimePicker1.Visible = False
            '    DateTimePicker3.Visible = True
            '    DateTimePicker2.Visible = True
            '    Button2.Visible = False
            '    ComboBox4.Visible = False
            '    ComboBox8.Visible = False
            'End If
            If ComboBox1.Text = "الشهر" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If

            If ComboBox1.Text = "الشهر لواردات الموردين" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "الشهر لواردات العملاء" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = False
            End If
            If ComboBox1.Text = "أسم الصنف" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = False
                ComboBox9.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = True
            End If
            If ComboBox1.Text = "أسم الصنف والشهر" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = True
                ComboBox9.Text = ""
            End If
            If ComboBox1.Text = "أسم الصنف ومورد" Then
                ComboBox2.Visible = True
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = False
                ComboBox2.Text = ""
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = False
                ComboBox9.Visible = True
                ComboBox9.Text = ""
            End If
            If ComboBox1.Text = "أسم الصنف وعميل" Then
                ComboBox2.Visible = False
                ComboBox6.Visible = False
                Button2.Visible = False
                ComboBox7.Visible = False
                ComboBox4.Visible = False

                ComboBox3.Visible = False
                ComboBox5.Visible = False
                DateTimePicker1.Visible = False
                DateTimePicker3.Visible = False
                DateTimePicker2.Visible = False
                ComboBox8.Visible = True
                ComboBox9.Visible = True
                ComboBox9.Text = ""
                ComboBox8.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox4_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox4.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox4.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub



    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox4.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub ComboBox6_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox6.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox6.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox6.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub ComboBox7_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Vendores Where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox7.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox7.Items.Add(rs_Vendors("Vend_Name").Value)
                rs_Vendors.MoveNext()
            Loop
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox7.SelectedItem.ToString()
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            vend_id = rs_Vendors("Vend_ID").Value
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If pr = "AllVend" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_All
                Dim frm_rpt As New Frm_rpt_ShowVend_All
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID And vend.Vend_visable='" & ssd & "') as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm, Vendores as vd Where sm.Vend_ID = vd.Vend_ID And vd.Vend_visable='" & ssd & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm, Vendores as vend Where sm.Vend_ID = vend.Vend_ID And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' And vend.Vend_visable='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)
                Dim gj As String
                gj = "الواردات للموردين"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", gj)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل أذونات الواردات للموردين"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "AllCust" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowCust_All
                Dim frm_rpt As New Frm_rpt_ShowCust_All
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID And vend.Cust_visable='" & ssd & "') as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm, Customers as vd Where sm.Cust_ID = vd.Cust_ID And vd.Cust_visable='" & ssd & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm, Customers as vend Where sm.Cust_ID = vend.Cust_ID And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' And vend.Cust_visable='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)
                Dim gj As String
                gj = "الواردات للعملاء"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", gj)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل أذونات الواردات للعملاء"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "fatora_Vend" Then
                'dy = DateTimePicker3.Value.Year
                'dm = DateTimePicker3.Value.Month
                'dd = DateTimePicker3.Value.Day
                'dstri = dy + "-" + dm + "-" + dd
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_fatora
                Dim frm_rpt As New Frm_rpt_ShowVend_fatora
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID =" & vend_id22 & ") as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Vend_ID =" & vend_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Total_Price) as c_No FROM Stor as sm Where sm.Vend_ID =" & vend_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_fato = 0
                Else
                    T_fato = rs_StoreVend("c_No").Value
                End If

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Payment) as c_No FROM Stor as sm Where sm.Vend_ID =" & vend_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    V_fato = 0
                Else
                    V_fato = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim ghf As Integer
                Dim sy As String
                sy = "الواردات"
                Dim sy3 As String
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Vendores Where Vend_ID=" & vend_id22 & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ghf = rs_StoreVend("Vend_Blance_Dein").Value
                sy3 = rs_StoreVend("Vend_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_fato)
                rpt.SetParameterValue("T_Name1", sy)
                rpt.SetParameterValue("T_Name2", sy3)
                rpt.SetParameterValue("p_name", V_fato)
                rpt.SetParameterValue("C_name", C_Name)
                rpt.SetParameterValue("B_name", ghf)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة فاتورة واردات لمورد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "fatora_Cust" Then
                'dy = DateTimePicker3.Value.Year
                'dm = DateTimePicker3.Value.Month
                'dd = DateTimePicker3.Value.Day
                'dstri = dy + "-" + dm + "-" + dd
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowCust_fatora
                Dim frm_rpt As New Frm_rpt_ShowCust_fatora
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID =" & cust_id22 & ") as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Cust_ID =" & cust_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Total_Price) as c_No FROM Stor as sm Where sm.Cust_ID =" & cust_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_fato = 0
                Else
                    T_fato = rs_StoreVend("c_No").Value
                End If

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Payment) as c_No FROM Stor as sm Where sm.Cust_ID =" & cust_id22 & " And sm.Stor_Datestri='" & dstri & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    V_fato = 0
                Else
                    V_fato = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim ghf As Integer
                Dim sy As String
                sy = "الواردات"
                Dim sy3 As String
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id22 & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ghf = rs_StoreVend("Cust_Blance_Medin").Value
                sy3 = rs_StoreVend("Cust_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_fato)
                rpt.SetParameterValue("T_Name1", sy)
                rpt.SetParameterValue("T_Name2", sy3)
                rpt.SetParameterValue("p_name", V_fato)
                rpt.SetParameterValue("C_name", C_Name)
                rpt.SetParameterValue("B_name", ghf)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة فاتورة واردات لعميل"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Vendor" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_Vendor
                Dim frm_rpt As New Frm_rpt_ShowVend_Vendor
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                'rcomand = New SqlCommand("Select (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Note From Stor as sm Where sm.Vend_ID=" & vend_id & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID =" & vend_id33 & ") as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Vend_ID=" & vend_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)

                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Vend_ID=" & vend_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim sy As String
                sy = "الواردات"
                Dim sy3 As String
                Dim s As Boolean
                s = False
                Dim ghf As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Vendores Where Vend_ID=" & vend_id33 & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ghf = rs_StoreVend("Vend_Blance_Dein").Value
                sy3 = rs_StoreVend("Vend_Blance_Type").Value
                rs_StoreVend.Close()

                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sy)
                rpt.SetParameterValue("T_Name2", sy3)
                rpt.SetParameterValue("V_Name", V_Name)
                rpt.SetParameterValue("B_name", ghf)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل الواردات لمورد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "month" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_month
                Dim frm_rpt As New Frm_rpt_ShowVend_month
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                'rcomand = New SqlCommand("Select (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id & " and sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID =sm.Vend_ID) as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)

                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Month_ID=" & month_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()

                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("mname", mo)
                rpt.SetParameterValue("T_Name", T_Name)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة التوريدات لشهر واحد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "monthvend" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_month
                Dim frm_rpt As New Frm_rpt_ShowVend_month
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID And vend.Vend_visable='" & ssd & "') as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm, Vendores as vd Where sm.Vend_ID = vd.Vend_ID And sm.Month_ID=" & month_id33 & " And vd.Vend_visable='" & ssd & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm, Vendores as vend Where sm.Vend_ID = vend.Vend_ID And sm.Month_ID=" & month_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' And vend.Vend_visable='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)
                Dim gj As String
                gj = "الواردات للموردين"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("mname", mo)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", gj)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الواردات للموردين لشهر ما"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "monthcust" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowCust_month
                Dim frm_rpt As New Frm_rpt_ShowCust_month
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID And vend.Cust_visable='" & ssd & "') as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm, Customers as vd Where sm.Cust_ID = vd.Cust_ID And sm.Month_ID=" & month_id33 & " And vd.Cust_visable='" & ssd & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm, Customers as vend Where sm.Cust_ID = vend.Cust_ID And sm.Month_ID=" & month_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' And vend.Cust_visable='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)
                Dim gj As String
                gj = "الواردات للعملاء"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("mname", mo)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", gj)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الواردات للموردين لشهر ما"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Vendor_month" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_month_Vendor
                Dim frm_rpt As New Frm_rpt_ShowVend_month_Vendor
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                'rcomand = New SqlCommand("Select (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id & " and sm.Vend_ID=" & vend_id & " and sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID =" & vend_id44 & ") as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id22 & " and sm.Vend_ID=" & vend_id44 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)

                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Month_ID=" & month_id22 & " and sm.Vend_ID=" & vend_id44 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()

                rada = New SqlDataAdapter(rcomand)
                Dim kl As String
                kl = "الواردات"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("mname", mo)
                rpt.SetParameterValue("V_Name", V_Name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", kl)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة واردات مورد فى شهر واحد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Customer_month" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowCust_month_Customer
                Dim frm_rpt As New Frm_rpt_ShowCust_month_Customer
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                'rcomand = New SqlCommand("Select (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id & " and sm.Vend_ID=" & vend_id & " and sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID =" & cust_id44 & ") as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Month_ID=" & month_id22 & " and sm.Cust_ID=" & cust_id44 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)

                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Month_ID=" & month_id22 & " and sm.Cust_ID=" & cust_id44 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()

                rada = New SqlDataAdapter(rcomand)
                Dim kl As String
                kl = "الواردات"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("mname", mo)
                rpt.SetParameterValue("V_Name", V_Name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", kl)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة واردات عميل فى شهر واحد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "date" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_date
                Dim frm_rpt As New Frm_rpt_ShowVend_date
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where  sm.Stor_Datestri='" & dstri2 & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Stor_Datestri='" & dstri2 & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الوارد ليوم"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Customer" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowCust_Customer
                Dim frm_rpt As New Frm_rpt_ShowCust_Customer
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                'rcomand = New SqlCommand("Select (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID) as vendname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Note From Stor as sm Where sm.Vend_ID=" & vend_id & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID =" & cust_id33 & ") as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Cust_ID=" & cust_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)

                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm Where sm.Cust_ID=" & cust_id33 & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim sy As String
                sy = "الواردات"
                Dim sy3 As String
                Dim s As Boolean
                s = False
                Dim ghf As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id33 & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ghf = rs_StoreVend("Cust_Blance_Medin").Value
                sy3 = rs_StoreVend("Cust_Blance_Type").Value
                rs_StoreVend.Close()

                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sy)
                rpt.SetParameterValue("T_Name2", sy3)
                rpt.SetParameterValue("V_Name", V_Name)
                rpt.SetParameterValue("B_name", ghf)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل الواردات لعميل"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "item" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowVend_All
                Dim frm_rpt As New Frm_rpt_ShowVend_All
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Vend"
                rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Vend_Name From Vendores as vend Where vend.Vend_ID = sm.Vend_ID And vend.Vend_visable='" & ssd & "') as vendname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm, Vendores as vd Where sm.Vend_ID = vd.Vend_ID And vd.Vend_visable='" & ssd & "' And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(sm.Stor_Quntity_No) as c_No FROM Stor as sm, Vendores as vend Where sm.Vend_ID = vend.Vend_ID And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' And vend.Vend_visable='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)
                Dim gj As String
                gj = "الواردات للموردين"
                rada.Fill(rdat, "Stor")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", gj)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل أذونات الواردات للموردين"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox8_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox8.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Customers Where Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox8.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox8.Items.Add(rs_Vendors("Cust_Name").Value)
                rs_Vendors.MoveNext()
            Loop
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox8.SelectedItem.ToString()
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            cust_id = rs_Vendors("Cust_ID").Value
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox9_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox9.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Itemes Where Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox9.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox9.Items.Add(rs_Vendors("Item_Name").Value)
                rs_Vendors.MoveNext()
            Loop
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox9.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox9.SelectedItem.ToString()
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Itemes Where Item_Name='" & u_name & "' And Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ite_id = rs_Vendors("Item_ID").Value

            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class