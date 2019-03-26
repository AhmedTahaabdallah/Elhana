Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class Frm_ShowImportsMoney

    Public month_id As Integer
    Public month_id8 As Integer
    Public month_id_re As Integer
    Public tp As String
    Public cust_id As Integer
    Public vend_id As Integer
    Public cust_id6 As Integer
    Public pr As String
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet
    Public T_Name As String
    Public M_Name As String
    Public tx_reson As String
    Public Cust_blance As Integer
    Public mr As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    'Public gcon As New OleDbConnection

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text = "كل أذونات الأيداع" Then
                ComboBox3.Visible = False
                Button1.Visible = True
                Button2.Visible = False
                TextBox1.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox4.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
            End If
            If ComboBox1.Text = "التاريخ" Then
                ComboBox3.Visible = False
                Button1.Visible = True
                Button2.Visible = False
                TextBox1.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox4.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = True
                ComboBox5.Visible = False
                ComboBox6.Visible = False
            End If
            If ComboBox1.Text = "أسم العميل" Then
                ComboBox3.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                Button1.Visible = True
                TextBox1.Visible = False
                Button2.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
            End If
            If ComboBox1.Text = "أسم العميل والشهر" Then
                ComboBox3.Visible = False
                ComboBox4.Visible = True
                ComboBox4.Text = ""
                ComboBox6.Visible = True
                ComboBox6.Text = ""
                Button1.Visible = True
                TextBox1.Visible = False
                Button2.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox5.Visible = False
            End If
            If ComboBox1.Text = "أسم المورد" Then
                ComboBox3.Visible = False
                ComboBox4.Visible = False
                ComboBox5.Visible = True
                ComboBox5.Text = ""
                Button1.Visible = True
                TextBox1.Visible = False
                Button2.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox6.Visible = False
            End If
            If ComboBox1.Text = "أسم المورد والشهر" Then
                ComboBox3.Visible = False
                ComboBox4.Visible = False
                Button1.Visible = True
                TextBox1.Visible = False
                Button2.Visible = False
                RadioButton1.Visible = False
                RadioButton2.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox5.Visible = True
                ComboBox5.Text = ""
                ComboBox6.Visible = True
                ComboBox6.Text = ""
            End If
            If ComboBox1.Text = "الشهر" Then
                ComboBox3.Visible = True
                ComboBox3.Text = ""
                ComboBox5.Visible = False
                Button1.Visible = True
                TextBox1.Visible = False
                RadioButton1.Visible = False
                Button2.Visible = False
                RadioButton2.Visible = False
                ComboBox4.Visible = False
                ComboBox2.Visible = False
                DateTimePicker1.Visible = False
                ComboBox6.Visible = False
            End If
            If ComboBox1.Text = "بحث متقدم" Then
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox2.Visible = False
                Button2.Visible = False
                Button1.Visible = False
                ComboBox4.Visible = False
                TextBox1.Visible = True
                RadioButton1.Visible = True
                RadioButton2.Visible = True
                RadioButton1.Checked = True
                DateTimePicker1.Visible = False
                ComboBox6.Visible = False
                TextBox1.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox3.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox3.Items.Add(rs_Cars("Month_Name").Value)
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
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim uh As Boolean
            uh = True
            If ComboBox1.Text = "كل أذونات الأيداع" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_all()
                'T_Name = TextBox2.Text

                If auth_report = uh Then
                    pr = "All"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    pr = "All"
                    Button2.Visible = True
                End If

            End If
            If ComboBox1.Text = "التاريخ" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_date()
                'T_Name = TextBox2.Text
                If auth_report = uh Then
                    pr = "date"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    pr = "date"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "الشهر" Then
                If ComboBox3.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox3.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_monthname()
                If auth_report = uh Then
                    month_id8 = month_id
                    'T_Name = TextBox2.Text
                    pr = "month"
                    M_Name = ComboBox3.Text
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    month_id8 = month_id
                    'T_Name = TextBox2.Text
                    pr = "month"
                    M_Name = ComboBox3.Text
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم العميل" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_Custname()
                If auth_report = uh Then
                    'T_Name = TextBox2.Text
                    cust_id6 = cust_id
                    M_Name = ComboBox4.Text
                    pr = "Customer"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    'T_Name = TextBox2.Text
                    cust_id6 = cust_id
                    M_Name = ComboBox4.Text
                    pr = "Customer"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم العميل والشهر" Then
                If ComboBox4.Text = "" Then
                    MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                    ComboBox4.Select()
                    Exit Sub
                End If
                If ComboBox6.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox6.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_Custname_cust()
                If auth_report = uh Then
                    'T_Name = TextBox2.Text
                    cust_id6 = cust_id
                    month_id8 = month_id
                    M_Name = ComboBox4.Text
                    pr = "month_Customer"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    'T_Name = TextBox2.Text
                    cust_id6 = cust_id
                    month_id8 = month_id
                    M_Name = ComboBox4.Text
                    pr = "month_Customer"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم المورد" Then
                If ComboBox5.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox5.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_Vendname()
                If auth_report = uh Then
                    'T_Name = TextBox2.Text
                    cust_id6 = vend_id

                    M_Name = ComboBox5.Text
                    pr = "Vendeor"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    'T_Name = TextBox2.Text
                    cust_id6 = vend_id

                    M_Name = ComboBox5.Text
                    pr = "Vendeor"
                    Button2.Visible = True
                End If
            End If
            If ComboBox1.Text = "أسم المورد والشهر" Then
                If ComboBox5.Text = "" Then
                    MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                    ComboBox5.Select()
                    Exit Sub
                End If
                If ComboBox6.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox6.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_Vendname_vend()
                If auth_report = uh Then
                    'T_Name = TextBox2.Text
                    cust_id6 = vend_id
                    month_id8 = month_id
                    M_Name = ComboBox5.Text
                    pr = "month_Vendeor"
                    Button2.Visible = True
                End If
                If u_Id = 1 Then
                    'T_Name = TextBox2.Text
                    cust_id6 = vend_id
                    month_id8 = month_id
                    M_Name = ComboBox5.Text
                    pr = "month_Vendeor"
                    Button2.Visible = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Frm_ShowImportsMoney_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            auth_report = rs_auth("User_moneyimp_repot").Value
            'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\SQLEXPRESS"
            If gcon.State = 1 Then gcon.Close()
            gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"
            gcon.Open()

            get_all()
            pr = "All"
            T_Name = TextBox2.Text
            Button2.Visible = True
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
            DataGridView1.Columns(2).Width = 100
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 350
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 243
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            'If u_Id = 1 Then
            '    DataGridView1.Columns(6).Width = 150
            '    DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            '    DataGridView1.Columns(7).Width = 160
            '    DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            'End If

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
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_date()
        Try
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Money_Datestri='" & dstri & "' And ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Money_Datestri='" & dstri & "' And ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Money_Datestri='" & dstri & "' And ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Money_Datestri='" & dstri & "' And ie.Month_ID=mnn.Month_ID And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_monthname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_Custname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_Custname_cust()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Cust_ID=" & cust_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_Vendname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub get_Vendname_vend()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim order_Type As String
            order_Type = "Imp"
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=ie.User_ID_edit) as [مستخدم التعديل] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            Else
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox2.Text = 0
            Else
                TextBox2.Text = rs_StoreVend("c_No").Value
            End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            rs_StoreVend.Close()


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id & " And mnn.Month_ID=" & month_id & " And ie.Vend_ID=" & vend_id & " And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag='" & ssd & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox2.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
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
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id_re = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        ComboBox2.Visible = False
        TextBox1.Text = ""
        tp = "re"
        Button2.Visible = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        ComboBox2.Visible = True
        TextBox1.Text = ""
        tp = "remo"
        Button2.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            If tp = "re" Then
                Dim gh As String
                gh = TextBox1.Text
                Dim ds As New DataSet
                Dim dt As New DataTable
                ds.Tables.Add(dt)
                Dim da As New OleDbDataAdapter
                Dim ssd As Boolean
                ssd = False
                Dim order_Type As String
                order_Type = "Imp"
                Dim stt As String
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "' ORDER BY ie.Money_Date DESC"
                da = New OleDbDataAdapter(stt, gcon)

                da.Fill(dt)
                DataGridView1.DataSource = dt.DefaultView
                gcon.Close()
                DataGridView1.Refresh()
                dgv()


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    TextBox2.Text = 0
                Else
                    TextBox2.Text = rs_StoreVend("c_No").Value
                End If
                'TextBox2.Text = rs_StoreVend("c_No").Value
                rs_StoreVend.Close()


                If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
                rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=mnn.Month_ID And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If IsDBNull(rs_StoreCust("c_No").Value) Then
                    Label3.Text = 0
                Else
                    Label3.Text = rs_StoreCust("c_No").Value
                End If
                rs_StoreCust.Close()
                'T_Name = TextBox2.Text
                pr = "re"
                tx_reson = gh
                Button2.Visible = True
            End If

            If tp = "remo" Then

                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If

                Dim gh As String
                gh = TextBox1.Text
                Dim ds As New DataSet
                Dim dt As New DataTable
                ds.Tables.Add(dt)
                Dim da As New OleDbDataAdapter
                Dim ssd As Boolean
                ssd = False
                Dim order_Type As String
                order_Type = "Imp"
                Dim stt As String
                stt = "Select ie.Money_ID as [رقم الأذن], ie.Money_Date as [التاريخ], mnn.Month_Name as [شهر], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id_re & " And mnn.Month_ID=" & month_id_re & " And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "' ORDER BY ie.Money_Date DESC"
                da = New OleDbDataAdapter(stt, gcon)

                da.Fill(dt)
                DataGridView1.DataSource = dt.DefaultView
                gcon.Close()
                DataGridView1.Refresh()
                dgv()


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("SELECT SUM(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id_re & " And mnn.Month_ID=" & month_id_re & " And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    TextBox2.Text = 0
                Else
                    TextBox2.Text = rs_StoreVend("c_No").Value
                End If
                'TextBox2.Text = rs_StoreVend("c_No").Value
                rs_StoreVend.Close()


                If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
                rs_StoreCust.Open("SELECT COUNT(ie.Money_ID) as c_No FROM Imp_Exp_Money as ie, year_monthes as mnn where ie.Month_ID=" & month_id_re & " And mnn.Month_ID=" & month_id_re & " And ie.Money_Reason Like '%" + gh + "%' And ie.Money_Type='" & order_Type & "' And mnn.Month_DeleFlag ='" & ssd & "' And ie.Money_DelFlag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If IsDBNull(rs_StoreCust("c_No").Value) Then
                    Label3.Text = 0
                Else
                    Label3.Text = rs_StoreCust("c_No").Value
                End If
                rs_StoreCust.Close()

                'T_Name = TextBox2.Text
                pr = "remo"
                mr = month_id_re
                tx_reson = gh
                M_Name = ComboBox2.Text
                Button2.Visible = True
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
            rs_Cars.Open("Select * From Customers Where Cust_DelFlag='" & s & "'  And Cust_visable='" & s & "' ORDER BY Cust_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox4.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox4.Items.Add(rs_Cars("Cust_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
            'End If
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
            rs_Cars.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            cust_id = rs_Cars("Cust_ID").Value
            'Cust_blance = rs_Cars("Cust_Blance_Medin").Value
            rs_Cars.Close()
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If pr = "All" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_All
                Dim frm_rpt As New Frm_rpt_ShowImp_All
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)

                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة كل أذونات الأيداع"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "date" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_Date
                Dim frm_rpt As New Frm_rpt_ShowImp_Date
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Money_Datestri='" & dstri & "' And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Money_Datestri='" & dstri & "' And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)

                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع بتاريخ ما"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "month" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_month
                Dim frm_rpt As New Frm_rpt_ShowImp_month
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("M_Name", M_Name)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع لشهر واحد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Customer" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_cust
                Dim frm_rpt As New Frm_rpt_ShowImp_cust
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Cust_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Cust_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim sio As String
                Dim kl As String
                Dim custty As String
                sio = "لعميل"
                kl = "أسم العميل "
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id6 & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Cust_blance = rs_StoreVend("Cust_Blance_Medin").Value
                custty = rs_StoreVend("Cust_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sio)
                rpt.SetParameterValue("T_Name2", kl)
                rpt.SetParameterValue("T_Name3", custty)
                rpt.SetParameterValue("M_Name", M_Name)
                rpt.SetParameterValue("B_Name", Cust_blance)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع لعميل"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "month_Customer" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_cust
                Dim frm_rpt As New Frm_rpt_ShowImp_cust
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Cust_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Cust_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim sio As String
                Dim kl As String
                Dim custty As String
                sio = "لعميل"
                kl = "أسم العميل "
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id6 & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Cust_blance = rs_StoreVend("Cust_Blance_Medin").Value
                custty = rs_StoreVend("Cust_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sio)
                rpt.SetParameterValue("T_Name2", kl)
                rpt.SetParameterValue("T_Name3", custty)
                rpt.SetParameterValue("M_Name", M_Name)
                rpt.SetParameterValue("B_Name", Cust_blance)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع لعميل"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "Vendeor" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_cust
                Dim frm_rpt As New Frm_rpt_ShowImp_cust
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Vend_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Vend_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim Vend_blance As Integer
                Dim sio As String
                Dim kl As String
                Dim vendty As String
                sio = "لمورد"
                kl = "أسم المورد "
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Vendores Where Vend_ID=" & cust_id6 & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Vend_blance = rs_StoreVend("Vend_Blance_Dein").Value
                vendty = rs_StoreVend("Vend_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sio)
                rpt.SetParameterValue("T_Name2", kl)
                rpt.SetParameterValue("T_Name3", vendty)
                rpt.SetParameterValue("M_Name", M_Name)
                rpt.SetParameterValue("B_Name", Vend_blance)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع لمورد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "month_Vendeor" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_cust
                Dim frm_rpt As New Frm_rpt_ShowImp_cust
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Vend_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Month_ID=" & month_id8 & " And ie.Vend_ID=" & cust_id6 & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                Dim s As Boolean
                s = False
                Dim Vend_blance As Integer
                Dim sio As String
                Dim kl As String
                Dim vendty As String
                sio = "لمورد"
                kl = "أسم المورد "
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Vendores Where Vend_ID=" & cust_id6 & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Vend_blance = rs_StoreVend("Vend_Blance_Dein").Value
                vendty = rs_StoreVend("Vend_Blance_Type").Value
                rs_StoreVend.Close()
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("T_Name1", sio)
                rpt.SetParameterValue("T_Name2", kl)
                rpt.SetParameterValue("T_Name3", vendty)
                rpt.SetParameterValue("M_Name", M_Name)
                rpt.SetParameterValue("B_Name", Vend_blance)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع لمورد"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "re" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_reason
                Dim frm_rpt As New Frm_rpt_ShowImp_reason
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Money_Reason Like '%" + tx_reson + "%' And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Money_Reason Like '%" + tx_reson + "%' And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)

                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع حسب البيان"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If

            If pr = "remo" Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New Rpt_ShowImp_reason_Month
                Dim frm_rpt As New Frm_rpt_ShowImp_reason_Month
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim tp As String
                tp = "Imp"
                'rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID = sm.Cust_ID) as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
                rcomand = New SqlCommand("Select ie.Money_ID, ie.Money_Date, ie.Money_Price, ie.Money_Reason, ie.Money_Note FROM Imp_Exp_Money as ie where ie.Money_Reason Like '%" + tx_reson + "%' And ie.Month_ID=" & mr & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date DESC", rcon)

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Sum(ie.Money_Price) as c_No FROM Imp_Exp_Money as ie where ie.Money_Reason Like '%" + tx_reson + "%' And ie.Month_ID=" & mr & " And ie.Money_Type='" & tp & "' And ie.Money_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If IsDBNull(rs_StoreVend("c_No").Value) Then
                    T_Name = 0
                Else
                    T_Name = rs_StoreVend("c_No").Value
                End If
                rada = New SqlDataAdapter(rcomand)

                rada.Fill(rdat, "Imp_Exp_Money")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("T_Name", T_Name)
                rpt.SetParameterValue("M_Name", M_Name)
                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة أذونات الأيداع حسب البيان لشهر واحد"
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

    Private Sub ComboBox5_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From Vendores Where Vend_DelFlag='" & s & "'  And Vend_visable='" & s & "' ORDER BY Vend_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox5.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox5.Items.Add(rs_Cars("Vend_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
            'End If
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
            rs_Cars.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            vend_id = rs_Cars("Vend_ID").Value
            'Cust_blance = rs_Cars("Cust_Blance_Medin").Value
            rs_Cars.Close()
            'End If
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
End Class