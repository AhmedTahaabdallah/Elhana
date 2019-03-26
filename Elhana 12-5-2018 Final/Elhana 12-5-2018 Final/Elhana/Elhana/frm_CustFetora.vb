Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class frm_CustFetora
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet
    Public oe As Integer
    Public st_request As String
    Public cust_id As Integer
    Public cust_idsearch As Integer
    Public cust_state_defult As String
    Public vend_idsearch As Integer
    Public vend_state_defult As String
    Public cust_iddelete As Integer
    Public vend_iddelete As Integer
    Public car_id As Integer
    Public Dri_Name As String
    Public gov_Name As String
    Public val_Old As Integer
    Public val_Oldtotlprice As Integer
    Public val_Oldpayment As Integer
    Public month_id As Integer
    Public ite_id As Integer
    Public store_quntity As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    Public button_Events As String
    Public cv As String
    Public isdafsave As String
    Public isdafedit As String
    Public isdafdelete As String
    Public order_Typesave As String
    Public order_Typeedit As String
    Public money_Typeedit As String
    Public stredit As String
    Public strnoteedit As String
    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 70
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 130
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
            DataGridView1.Columns(8).Width = 200
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub getallfetora()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String
            stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And st.Fetora_id=" & Fetora_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_ID DESC"
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT SUM(st.Stor_Total_Price) as c_No FROM Stor as st, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And st.Fetora_id=" & Fetora_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label6.Text = 0
            Else
                Label6.Text = rs_StoreVend("c_No").Value
            End If

            rs_StoreVend.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT SUM(st.Stor_Payment) as c_No FROM Stor as st, Itemes as itm, year_monthes as mn where st.Cust_ID=" & cust_id & " And st.Fetora_id=" & Fetora_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label32.Text = 0
            Else
                Label32.Text = rs_StoreCust("c_No").Value
            End If

            Dim f_totalprice As Integer
            Dim f_totalpay As Integer
            f_totalprice = Label6.Text
            f_totalpay = Label32.Text
            Dim c_blance As Integer
            Dim c_type As String
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID='" & cust_id & "' And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            c_type = rs_Vendors("Cust_Blance_Type").Value
            c_blance = rs_Vendors("Cust_Blance_Medin").Value
            s = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Fetora Where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_StoreVend("Fetora_totalprice").Value = f_totalprice
            rs_StoreVend("Fetora_totalpay").Value = f_totalpay
            If st_request = "notload" Then
                rs_StoreVend("Fetora_blanceafter").Value = c_blance
                rs_StoreVend("Fetora_blanceaftertype").Value = c_type
            End If
            
            rs_StoreVend.Update()

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Fetora Where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Label38.Text = rs_StoreVend("Fetora_blancebeforetype").Value
            Label2.Text = rs_StoreVend("Fetora_blancebefore").Value
            Label6.Text = rs_StoreVend("Fetora_totalprice").Value
            Label32.Text = rs_StoreVend("Fetora_totalpay").Value
            Label39.Text = rs_StoreVend("Fetora_blanceaftertype").Value
            Label31.Text = rs_StoreVend("Fetora_blanceafter").Value
           
            Dim s1 As Boolean
            s1 = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'ComboBox2.Text = rs_Vendors("Cust_Name").Value
            Label40.Text = rs_Vendors("Cust_Blance_Type").Value
            'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            Label41.Text = rs_Vendors("Cust_Blance_Medin").Value
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub frm_CustFetora_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_storcust_add").Value
            auth_show = rs_auth("User_storcust_show").Value
            auth_search = rs_auth("User_storcust_search").Value
            auth_edit = rs_auth("User_storcust_edit").Value
            auth_delete = rs_auth("User_storcust_delete").Value

            If auth_add = ut Then
                'btn_add.Visible = False
                btn_save.Visible = False
            Else
                'btn_add.Visible = True
                btn_save.Visible = True
            End If

            If auth_search = ut Then
                Label7.Visible = False
                TextBox3.Visible = False
                Button1.Visible = False
            Else
                Label7.Visible = True
                TextBox3.Visible = True
                Button1.Visible = True
            End If
            If auth_edit = ut Then
                btn_etite.Visible = False
            Else
                btn_etite.Visible = True
            End If
            If auth_delete = ut Then
                btn_delete.Visible = False
            Else
                btn_delete.Visible = True
            End If
            auth_stor = rs_auth("User_stor").Value
            auth_blance = rs_auth("User_blance").Value

            If auth_stor = ut Then
                Label17.Visible = False
                Label15.Visible = False
                Label18.Visible = False
            Else
                Label17.Visible = True
                Label15.Visible = True
                Label18.Visible = True
            End If
            If auth_blance = ut Then
                Label29.Visible = False
                Label27.Visible = False
                Label28.Visible = False
            Else
                Label29.Visible = True
                Label27.Visible = True
                Label28.Visible = True
            End If
            If u_Id = 1 Then
                'btn_add.Visible = True
                btn_save.Visible = True
                Label7.Visible = True
                TextBox3.Visible = True
                Button1.Visible = True
                btn_etite.Visible = True
                btn_delete.Visible = True
                Label17.Visible = True
                Label15.Visible = True
                Label18.Visible = True
                Label29.Visible = True
                Label27.Visible = True
                Label28.Visible = True
            End If

            oe = 1


            Label1.Visible = True
            ComboBox2.Visible = True

            ComboBox7.Text = ""
            ComboBox2.Text = ""
            ComboBox1.Text = ""

            txt_quntity.Text = ""
            TextBox1.Text = ""
            Label1.Text = Fetora_id
            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Label27.Text = rs_Store("Blance_Total").Value
            End If


            Dim s As Boolean
            s = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Fetora Where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If fetore_type = "serch" Then
                DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                cust_id = rs_StoreVend("Fetora_custid").Value
                month_id = rs_StoreVend("Fetora_monthid").Value


            Else
                DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
                cust_id = Fetora_custid
                month_id = Fetora_monthid
            End If
            Label38.Text = rs_StoreVend("Fetora_blancebeforetype").Value
            Label2.Text = rs_StoreVend("Fetora_blancebefore").Value
            Label6.Text = rs_StoreVend("Fetora_totalprice").Value
            Label32.Text = rs_StoreVend("Fetora_totalpay").Value
            Label39.Text = rs_StoreVend("Fetora_blanceaftertype").Value
            Label31.Text = rs_StoreVend("Fetora_blanceafter").Value
            Dim s1 As Boolean
            s1 = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'ComboBox2.Text = rs_Vendors("Cust_Name").Value
            Label40.Text = rs_Vendors("Cust_Blance_Type").Value
            'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            Label41.Text = rs_Vendors("Cust_Blance_Medin").Value
            'cust_id = Fetora_custid

            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_ID=" & month_id & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Text = rs_Cars("Month_Name").Value
            'month_id = Fetora_monthid
            st_request = "load"
            getallfetora()




            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox7.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""

            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"

            btn_etite.Enabled = False
            btn_delete.Enabled = False
            btn_add.Enabled = False
            Label7.Enabled = False
            'TextBox3.Enabled = False
            'Button1.Enabled = False
            btn_save.Enabled = True

            ComboBox7.Enabled = True
            Button2.Visible = True
            ComboBox7.Select()
            'Dim s As Boolean
            s = False
            Try

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Count(Stor_ID) as df From Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                If rs_StoreVend("df").Value = 0 Then
                    Label11.Text = 1
                Else
                    Dim u_name As Integer
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    u_name = rs_StoreVend("Stor_ID").Value
                    Label11.Text = u_name + 1
                End If


                'Dim s1 As Boolean
                s1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & Fetora_custid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox2.Text = rs_Vendors("Cust_Name").Value
                Label19.Text = rs_Vendors("Cust_Blance_Type").Value
                'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
                Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
                cust_id = Fetora_custid

                If rs_Cars.State = 1 Then rs_Cars.Close()
                rs_Cars.Open("Select * From year_monthes Where Month_ID=" & Fetora_monthid & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                ComboBox1.Text = rs_Cars("Month_Name").Value
                month_id = Fetora_monthid

                If fetore_type = "serch" Then


                    Dim order_Type As String
                    order_Type = "Cust"
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Fetora where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                    End If
                    DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                    'rs_StoreVend.Close()
                Else
                    DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            End Try



            rs_Vendors.Close()
            rs_Cars.Close()
            rs_StoreVend.Close()
            rs_StoreCust.Close()
            'rs_StoreVend.Close()
            'rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub









    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_quntity.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub





    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click


        Label1.Visible = True
        ComboBox2.Visible = True
        ComboBox7.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""

        txt_quntity.Text = 0
        txt_quntity.Text = 0
        txt_Payment.Text = 0
        txt_price.Text = 0
        TextBox1.Text = ""
        Label15.Text = "-"
        Label16.Text = "-"

        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        Label7.Enabled = False
        TextBox3.Enabled = False
        Button1.Enabled = False
        btn_save.Enabled = True

        ComboBox7.Enabled = True
        Button2.Visible = True
        ComboBox7.Select()
        Dim s As Boolean
        s = False
        Try

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(Stor_ID) as df From Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label11.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("Stor_ID").Value
                Label11.Text = u_name + 1
            End If


            Dim s1 As Boolean
            s1 = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & Fetora_custid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox2.Text = rs_Vendors("Cust_Name").Value
            Label19.Text = rs_Vendors("Cust_Blance_Type").Value
            'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
            cust_id = Fetora_custid

            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_ID=" & Fetora_monthid & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Text = rs_Cars("Month_Name").Value
            month_id = Fetora_monthid

            If fetore_type = "serch" Then
               

                Dim order_Type As String
                order_Type = "Cust"
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Fetora where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                End If
                DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                'rs_StoreVend.Close()
            Else
                DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If ComboBox7.Text = "" Then
            MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
            ComboBox7.Select()
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If

        If txt_Payment.Text = "" Then
            txt_Payment.Text = 0
        End If

        'If ComboBox3.Text = "" Then
        '    MsgBox("من فضلك أختر رقم السيارة", MsgBoxStyle.Information, "تنبيه")
        '    Exit Sub
        'End If

        'If ComboBox5.Text = "" Then
        '    MsgBox("من فضلك أختر أسم السائق", MsgBoxStyle.Information, "تنبيه")
        '    Exit Sub
        'End If

        'If ComboBox4.Text = "" Then
        '    MsgBox("من فضلك أختر الجهة القادم منها الشحنة الواردة", MsgBoxStyle.Information, "تنبيه")
        '    Exit Sub
        'End If

        'If ComboBox6.Text = "" Then
        '    MsgBox("من فضلك أختر الجهة التى سوف يتم أستلام قيها الشحنة الواردة", MsgBoxStyle.Information, "تنبيه")
        '    Exit Sub
        'End If

        If txt_quntity.Text = "" Then
            MsgBox("من فضلك أدخل الكمية ", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_price.Text = "" Then
            MsgBox("من فضلك أدخل السعر ", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_totalPrice.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى السعر ", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        Dim jk As Integer

        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
            jk = 1
        End If
        Try
            Dim uk As Integer
            uk = 2
            Dim s As Boolean
            s = False
            Dim stor_In As Integer
            Dim stor_Out As Integer
            Dim quentity As Integer
            Dim val_storintotalprice As Integer
            'Dim val_pay As Integer
            quentity = Val(txt_quntity.Text)
            Dim toprice As Integer
            toprice = Val(txt_totalPrice.Text)
            Dim price As Integer
            price = Val(txt_price.Text)
            Dim payme As Integer
            payme = Val(txt_Payment.Text)
            Dim vendname As String

            order_Typesave = "Cust"
            vendname = ComboBox2.Text
            'Dim u_name As Integer
            'Dim u_name1 As Integer
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAX(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'u_name = rs_StoreVend("Stor_ID").Value
            'u_name1 = u_name + 1
            Dim z As Integer
            z = 0

            ' '' '' ''To sure about payment not make blance of Customers < 0
            Dim tt As Integer
            tt = cust_id
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                isdafsave = rs_Vendors("Cust_Blance_d").Value
            End If
            'val_pay = txt_Payment.Text
            'Dim ty44 As Integer
            'ty44 = val_storintotalprice + toprice
            'Dim ty2 As Integer
            'ty2 = ty44 - txt_Payment.Text

            'If ty2 < z Then
            '    MsgBox("رصيد العميل يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
            '    txt_quntity.Select()
            '    Exit Sub
            'End If

            stor_Out = txt_quntity.Text
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            stor_In = rs_Store("Store_Quntity").Value

            If stor_In < stor_Out Then
                MsgBox("الكمية داخل المخزن لا تكفى لصرف الاذن", MsgBoxStyle.Information, "تنبيه")
                txt_quntity.Select()
                Exit Sub
            End If
            '' '' '' ''--------------------------------------------
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            s = False
            Dim sg As String
            sg = "رصيد مبدئى " + val_storintotalprice.ToString

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Customers as vd where st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Cust_ID=vd.Cust_ID And st.Stor_DelFlag='" & s & "' And vd.Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreCust("c_No").Value) Then
               
                
            Else
                Dim d As Integer
                d = rs_StoreCust("c_No").Value
                If d = 0 Then
                    If jk = 1 Then
                        TextBox1.Text = sg
                    Else
                        TextBox1.Text = TextBox1.Text + sg
                    End If
                End If
            End If
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_StoreVend.AddNew()
            'rs_StoreVend("Stor_ID").Value = u_name1
            'rs_StoreVend("car_Id").Value = car_id
            'rs_StoreVend("Stor_From").Value = gov_Name
            'rs_StoreVend("Stor_To").Value = ComboBox6.Text
            'rs_StoreVend("Drivers_Name").Value = Dri_Name
            rs_StoreVend("User_ID").Value = u_Id
            rs_StoreVend("User_ID_edit").Value = uk
            rs_StoreVend("User_ID_delete").Value = uk
            rs_StoreVend("Cust_ID").Value = cust_id
            rs_StoreVend("Vend_ID").Value = oe
            rs_StoreVend("Stor_Date").Value = DateTimePicker1.Value.ToShortDateString()
            rs_StoreVend("Stor_Datestri").Value = dstri
            rs_StoreVend("Stor_Quntity_No").Value = quentity
            rs_StoreVend("Stor_DelFlag").Value = s
            rs_StoreVend("Item_ID").Value = ite_id
            rs_StoreVend("Month_ID").Value = month_id
            rs_StoreVend("Stor_Note").Value = TextBox1.Text
            rs_StoreVend("Stor_Type").Value = order_Typesave
            rs_StoreVend("Stor_Price_tin").Value = price
            rs_StoreVend("Stor_Total_Price").Value = toprice
            rs_StoreVend("Stor_Payment").Value = payme
            rs_StoreVend("Stor_Type_Arabic").Value = "صادر الى"
            rs_StoreVend("Fetora_id").Value = Fetora_id
            rs_StoreVend.Update()
            MsgBox("تم حفظ عملية الصادر بنجاح", MsgBoxStyle.Information, "حفظ بيانات الصادرات")
            ComboBox2.Text = ""
            ComboBox1.Text = ""

            ComboBox7.Text = ""
            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"

            btn_etite.Enabled = True
            btn_delete.Enabled = True
            btn_add.Enabled = True
            Label7.Enabled = True
            TextBox3.Enabled = True
            Button1.Enabled = True
            btn_save.Enabled = False
            Dim ne22 As Integer
            'ne22 = rs_StoreVend("Stor_Quntity_No").Value
            ne22 = quentity

            Dim s3 As Boolean
            s3 = False
            Dim ne As Integer
            Dim fin As Integer
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("Select * From Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ne = rs_Store("Store_Quntity").Value
            fin = ne - ne22

            Dim s8 As Boolean
            s8 = False
            Dim u_name8 As Integer
            Dim n_Type As String
            n_Type = "Cust"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select* From Stor Where Stor_Type ='" & n_Type & "' And Stor_DelFlag='" & s8 & "' And Stor_ID = (SELECT MAx(Stor_ID) FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            u_name8 = rs_StoreVend("Stor_ID").Value

            Dim mm1 As String
            mm1 = "حفظ صادر"
            s = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                rs_Store("Store_Quntity").Value = fin
                rs_Store("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store("Stor_ID").Value = u_name8
                rs_Store("Store_DelFlag").Value = s3
                rs_Store("User_ID").Value = u_Id
                rs_Store("Stroe_State").Value = mm1
                rs_Store("Item_ID").Value = ite_id
                rs_Store.Update()
            Else

                rs_Store("Store_Quntity").Value = fin
                rs_Store("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store("Stor_ID").Value = u_name8
                rs_Store("Store_DelFlag").Value = s3
                rs_Store("User_ID").Value = u_Id
                rs_Store("Stroe_State").Value = mm1
                rs_Store("Item_ID").Value = ite_id
                rs_Store.Update()
            End If




            ' ''عليه رصيد
            If isdafsave = "defulit" Then
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    Dim jj As Integer
                    jj = val_storintotalprice + toprice
                    rs_Vendors("Cust_Blance_Medin").Value = jj
                    rs_Vendors.Update()
                End If


            End If

            ' ''ليه رصيد
            If isdafsave = "notdefulit" Then
                If val_storintotalprice > toprice Then
                    Dim ff As Integer
                    ff = val_storintotalprice - toprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Cust_Blance_Medin").Value = ff
                        rs_Vendors.Update()
                    End If
                End If

                If val_storintotalprice <= toprice Then
                    Dim ff8 As Integer
                    ff8 = toprice - val_storintotalprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim re As String
                        Dim qw As String
                        re = "defulit"
                        qw = "عليه رصيد"
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Cust_Blance_Medin").Value = ff8
                        rs_Vendors("Cust_Blance_Type").Value = qw
                        rs_Vendors("Cust_Blance_d").Value = re
                        rs_Vendors.Update()
                    End If
                End If
            End If
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                isdafsave = rs_Vendors("Cust_Blance_d").Value
            End If

            If payme > 0 Then
                dy = DateTimePicker1.Value.Year
                dm = DateTimePicker1.Value.Month
                dd = DateTimePicker1.Value.Day
                dstri = dy + "-" + dm + "-" + dd
                order_Typesave = "Imp"
                Dim str As String
                str = "وارد من " + vendname
                s = False
                Dim strnote As String
                Dim strn2 As String
                strn2 = u_name8
                strnote = "رقم أذن الصادر للمخزن " + strn2
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Customers.AddNew()
                'rs_Customers("Money_ID").Value = u_name1
                rs_Customers("Money_Type").Value = order_Typesave
                rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
                rs_Customers("Money_Datestri").Value = dstri
                rs_Customers("Money_Price").Value = payme
                rs_Customers("Money_Reason").Value = str
                rs_Customers("Money_Note").Value = strnote
                rs_Customers("User_ID").Value = u_Id
                rs_Customers("Money_DelFlag").Value = s
                rs_Customers("Month_ID").Value = month_id
                rs_Customers("Stor_ID").Value = u_name8
                rs_Customers("Cust_ID").Value = cust_id
                rs_Customers("Vend_ID").Value = oe
                rs_Customers("Money_Type_Arabic").Value = "وارد إلى الخزنة"
                rs_Customers.Update()

                ' ''عليه رصيد
                If isdafsave = "defulit" Then

                    If val_storintotalprice >= payme Then
                        Dim ff As Integer
                        ff = val_storintotalprice - payme
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Cust_Blance_Medin").Value = ff
                            rs_Vendors.Update()
                        End If
                    End If

                    If val_storintotalprice < payme Then
                        Dim ff8 As Integer
                        ff8 = payme - val_storintotalprice
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "ليه رصيد"
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Cust_Blance_Medin").Value = ff8
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If

                End If

                ' ''ليه رصيد
                If isdafsave = "notdefulit" Then
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim jj As Integer
                        jj = val_storintotalprice + payme
                        rs_Vendors("Cust_Blance_Medin").Value = jj
                        rs_Vendors.Update()
                    End If

                End If


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_Type ='" & order_Typesave & "' And Money_DelFlag='" & s8 & "' And Money_ID = (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name8 = rs_StoreVend("Money_ID").Value

                'Dim s3 As Boolean
                s3 = False
                Dim ne9 As Integer
                Dim ne229 As Integer
                Dim fin9 As Integer
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ne9 = rs_Store("Blance_Total").Value
                ne229 = rs_StoreVend("Money_Price").Value
                fin9 = ne9 + ne229

                Dim mm13 As String
                mm13 = "حفظ وارد"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = u_name8
                rs_Store("Blance_Total").Value = fin9
                rs_Store("DelFlag").Value = s3
                rs_Store("Blance_State").Value = mm13
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()
                rs_Store.Close()
            End If

            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Label27.Text = rs_Store("Blance_Total").Value
            End If

            rs_Store.Close()

            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"
            Button2.Visible = False
            st_request = "notload"
            getallfetora()


            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox7.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""

            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"

            btn_etite.Enabled = False
            btn_delete.Enabled = False
            btn_add.Enabled = False
            Label7.Enabled = False
            'TextBox3.Enabled = False
            'Button1.Enabled = False
            btn_save.Enabled = True

            ComboBox7.Enabled = True
            Button2.Visible = True
            ComboBox7.Select()
            'Dim s As Boolean
            s = False
            Try

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Count(Stor_ID) as df From Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                If rs_StoreVend("df").Value = 0 Then
                    Label11.Text = 1
                Else
                    Dim u_name As Integer
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    u_name = rs_StoreVend("Stor_ID").Value
                    Label11.Text = u_name + 1
                End If


                Dim s1 As Boolean
                s1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & Fetora_custid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox2.Text = rs_Vendors("Cust_Name").Value
                Label19.Text = rs_Vendors("Cust_Blance_Type").Value
                'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
                Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
                cust_id = Fetora_custid

                If rs_Cars.State = 1 Then rs_Cars.Close()
                rs_Cars.Open("Select * From year_monthes Where Month_ID=" & Fetora_monthid & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                ComboBox1.Text = rs_Cars("Month_Name").Value
                month_id = Fetora_monthid

                If fetore_type = "serch" Then


                    Dim order_Type As String
                    order_Type = "Cust"
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Fetora where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                    End If
                    DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                    'rs_StoreVend.Close()
                Else
                    DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            End Try


            rs_StoreVend.Close()
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox3.Text = "0" Or TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            Dim u_name As Integer
            u_name = TextBox3.Text
            Dim order_Type As String
            order_Type = "Cust"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & s & "' And Stor_ID=" & u_name & " And Stor_Type='" & order_Type & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                TextBox3.Select()
                Exit Sub
            End If

            'ComboBox5.Text = rs_StoreVend("Drivers_Name").Value
            'Dri_Name = rs_StoreVend("Drivers_Name").Value
            TextBox1.Text = rs_StoreVend("Stor_Note").Value
            txt_quntity.Text = rs_StoreVend("Stor_Quntity_No").Value
            'ComboBox6.Text = rs_StoreVend("Stor_To").Value
            'ComboBox4.Text = rs_StoreVend("Stor_From").Value
            'gov_Name = rs_StoreVend("Stor_From").Value
            DateTimePicker1.Value = rs_StoreVend("Stor_Date").Value
            Label11.Text = rs_StoreVend("Stor_ID").Value
            If IsDBNull(rs_StoreVend("Fetora_id").Value) Then
                Label1.Text = 0
            Else
                Label1.Text = rs_StoreVend("Fetora_id").Value
            End If

            txt_price.Text = rs_StoreVend("Stor_Price_tin").Value
            txt_totalPrice.Text = rs_StoreVend("Stor_Total_Price").Value
            val_Oldtotlprice = rs_StoreVend("Stor_Total_Price").Value
            txt_Payment.Text = rs_StoreVend("Stor_Payment").Value
            val_Oldpayment = rs_StoreVend("Stor_Payment").Value
            val_Old = txt_quntity.Text
            Dim uio As Integer
            uio = rs_StoreVend("Cust_ID").Value

            'If uio = oe Then
            '    Dim s14 As Boolean
            '    s14 = False
            '    Dim gh4 As Integer
            '    gh4 = rs_StoreVend("Vend_ID").Value
            '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
            '    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & gh4 & " And Vend_DelFlag='" & s14 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    Label19.Text = rs_Vendors("Vend_Blance_Type").Value
            '    vend_state_defult = rs_Vendors("Vend_Blance_d").Value
            '    vend_idsearch = gh4
            '    vend_iddelete = gh4
            '    cv = "vend"
            '    Label16.Text = rs_Vendors("Vend_Blance_Dein").Value
            '    ComboBox2.Visible = False
            '    Label1.Visible = False

            '    Label31.Visible = True
            'End If
            'If uio > oe Then
            Dim s1 As Boolean
            s1 = False
            Dim gh As Integer
            gh = rs_StoreVend("Cust_ID").Value
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & gh & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox2.Text = rs_Vendors("Cust_Name").Value
            Label19.Text = rs_Vendors("Cust_Blance_Type").Value
            cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            cust_idsearch = gh
            cust_iddelete = gh
            cv = "cust"
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
            ComboBox2.Visible = True
            Label1.Visible = True
            'ComboBox2.Enabled = False

            'End If


            Dim s13 As Boolean
            s13 = False
            Dim gh1 As Integer
            Dim jh As String
            gh1 = rs_StoreVend("Item_ID").Value
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Itemes Where Item_ID=" & gh1 & " And Item_DelFlag='" & s13 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox7.Text = rs_Vendors("Item_Name").Value
            ite_id = gh1
            jh = rs_Vendors("Item_Wight_Type").Value
            If jh = "وحدة" Then
                Label30.Text = jh
                Label18.Text = jh
                Label21.Text = "سعر الوحدة :"
            End If
            If jh = "كيلو" Then
                Label30.Text = "طن"
                Label18.Text = "كيلو"
                Label21.Text = "سعر الطن :"
            End If
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Label15.Text = rs_Vendors("Store_Quntity").Value
            'store_quntity = rs_Vendors("Store_Quntity").Value
            Dim s15 As Boolean
            s15 = False
            Dim gh5 As Integer
            gh5 = rs_StoreVend("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.Open("Select * From year_monthes Where Month_ID=" & gh5 & " And Month_DeleFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox1.Text = rs_U("Month_Name").Value
            month_id = gh5

            'Dim s2 As Boolean
            's = False
            'Dim u_name1 As Integer
            'u_name1 = rs_StoreVend("car_Id").Value
            'If rs_Cars.State = 1 Then rs_Cars.Close()
            'rs_Cars.Open("Select * From Cars_Drivers Where car_Id=" & u_name1 & " And Car_DelFlag='" & s2 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            'ComboBox3.Text = rs_Cars("Car_No").Value
            'car_id = u_name1
            'rs_Cars.Close()
            TextBox3.Text = ""
            rs_Vendors.Close()
            rs_StoreVend.Close()
            rs_U.Close()

            ComboBox7.Enabled = False
            btn_save.Enabled = False

            btn_etite.Enabled = True
            btn_delete.Enabled = True
            button_Events = "edit"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        If Label11.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        If ComboBox7.Text = "" Then
            MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
            ComboBox7.Select()
            Exit Sub
        End If


        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If

        'If ComboBox3.Text = "" Then
        '    MsgBox("من فضلك أختر رقم السيارة")
        '    Exit Sub
        'End If

        'If ComboBox5.Text = "" Then
        '    MsgBox("من فضلك أختر أسم السائق")
        '    Exit Sub
        'End If

        'If ComboBox4.Text = "" Then
        '    MsgBox("من فضلك أختر الجهة القادم منها الشحنة الواردة")
        '    Exit Sub
        'End If

        'If ComboBox6.Text = "" Then
        '    MsgBox("من فضلك أختر الجهة التى سوف يتم أستلام قيها الشحنة الواردة")
        '    Exit Sub
        'End If

        If txt_quntity.Text = "" Then
            MsgBox("من فضلك أدخل الكمية ")
            Exit Sub
        End If
        If txt_price.Text = "" Then
            MsgBox("من فضلك أدخل السعر ")
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim val_storin As Integer
            Dim val_Blance_Total As Integer
            Dim x As Integer
            Dim val_addtostore As Integer
            Dim val_new As Integer
            Dim val_storintotalprice As Integer
            Dim xtotalprice As Integer
            Dim val_addtototalprice As Integer
            Dim val_newtotalprice As Integer
            Dim xpayment As Integer
            Dim val_addtopayment As Integer
            Dim val_newpayment As Integer
            Dim val_Blance_Total2 As Integer
            Dim rr As Integer
            Dim ss1 As Boolean
            ss1 = False
            money_Typeedit = "Imp"
            order_Typeedit = "Cust"
            'If cv = "cust" Then


            'End If
            'If cv = "vend" Then
            '    order_Typeedit = "Vend"

            'End If



            Dim f As Integer
            f = Label11.Text
            Dim f1 As String
            f1 = Label11.Text
            Dim quentity As Integer
            quentity = Val(txt_quntity.Text)
            Dim toprice As Integer
            toprice = Val(txt_totalPrice.Text)
            Dim price As Integer
            price = Val(txt_price.Text)
            Dim payme As Integer
            payme = Val(txt_Payment.Text)
            Dim z As Integer
            z = 0
            Dim tt As Integer

            ' '' '' ''To sure about totalprice and payment not make blance of Customers < 0
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                val_Blance_Total2 = rs_Store("Blance_Total").Value

            End If
            If cv = "cust" Then

                tt = cust_idsearch
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If


                'val_newtotalprice = toprice
                'If val_newtotalprice < val_Oldtotlprice Then
                '    val_addtototalprice = val_Oldtotlprice - val_newtotalprice
                '    Dim ll As Integer
                '    ll = val_storintotalprice - val_addtototalprice
                '    If ll < z Then
                '        MsgBox("رصيد العميل يجب الايكون أقل من الصفر  (قم بزيادة الكمية ليزيد المبلغ الإجمالى) ", MsgBoxStyle.Information, "تنبيه")
                '        txt_quntity.Select()
                '        Exit Sub
                '    End If
                'End If
                'Dim ade As Integer
                'Dim ades As Integer
                'val_newpayment = txt_Payment.Text
                'If val_Oldpayment < val_newpayment Then
                '    val_addtopayment = val_newpayment - val_Oldpayment
                '    If val_Oldtotlprice > toprice Then
                '        ade = val_Oldtotlprice - toprice
                '        ades = val_storintotalprice - ade
                '    End If
                '    If val_Oldtotlprice < toprice Then
                '        ade = toprice - val_Oldtotlprice
                '        ades = val_storintotalprice + ade
                '    End If
                '    If val_Oldtotlprice = toprice Then
                '        ades = val_storintotalprice
                '    End If
                '    Dim ty33 As Integer
                '    ty33 = ades - val_addtopayment
                '    If ty33 < z Then
                '        MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '        txt_quntity.Select()
                '        Exit Sub
                '    End If
                'End If

                val_newpayment = payme
                If val_Oldpayment > val_newpayment Then
                    val_addtopayment = val_Oldpayment - val_newpayment
                    Dim ty222 As Integer
                    ty222 = val_Blance_Total2 - val_addtopayment
                    If ty222 < z Then
                        MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر (قم بزيادة المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                        txt_Payment.Select()
                        Exit Sub
                    End If
                End If

            End If
            If cv = "vend" Then
                tt = vend_idsearch
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                    isdafedit = rs_Vendors("Vend_Blance_d").Value
                End If
                val_newpayment = payme
                If val_Oldpayment > val_newpayment Then
                    val_addtopayment = val_Oldpayment - val_newpayment
                    Dim ty222 As Integer
                    ty222 = val_Blance_Total2 - val_addtopayment
                    If ty222 < z Then
                        MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر (قم بزيادة المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                        txt_Payment.Select()
                        Exit Sub
                    End If
                End If
            End If




            ' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''For stor
            ss1 = False
            val_new = quentity
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            val_storin = rs_Store("Store_Quntity").Value

            If val_Old > val_new Then
                val_addtostore = val_Old - val_new
                x = val_storin + val_addtostore
            End If


            If val_new > val_Old Then
                val_addtostore = val_new - val_Old
                If val_storin >= val_addtostore Then
                    x = val_storin - val_addtostore
                End If
                'If val_storin = val_addtostore Then
                '    x = 0
                'End If
                If val_storin < val_addtostore Then
                    MsgBox("الكمية داخل المخزن لا تكفى للتعديل", MsgBoxStyle.Information, "تنبيه")
                    txt_quntity.Select()
                    Exit Sub
                End If
            End If
            If val_Old = val_new Then
                x = val_storin
            End If

            '' '' '' '' '' '' '' '' '' '' '' '' '' ''For  totalprice

            If cv = "cust" Then
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If

                val_newtotalprice = toprice
                If val_Oldtotlprice < val_newtotalprice Then
                    val_addtototalprice = val_newtotalprice - val_Oldtotlprice
                    If isdafedit = "defulit" Then
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                    If isdafedit = "notdefulit" Then

                        If val_storintotalprice > val_addtototalprice Then
                            xtotalprice = val_storintotalprice - val_addtototalprice
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice <= val_addtototalprice Then
                            xtotalprice = val_addtototalprice - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "defulit"
                            qw = "عليه رصيد"
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If

                    End If

                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else

                    'End If
                End If


                If val_newtotalprice < val_Oldtotlprice Then
                    val_addtototalprice = val_Oldtotlprice - val_newtotalprice
                    'Dim tow As Integer
                    'tow = val_storintotalprice - val_addtototalprice
                    'If tow < z Then
                    '    MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                    '    txt_quntity.Select()
                    '    Exit Sub
                    'End If
                    If isdafedit = "defulit" Then

                        If val_addtototalprice <= val_storintotalprice Then
                            xtotalprice = val_storintotalprice - val_addtototalprice
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_addtototalprice > val_storintotalprice Then
                            xtotalprice = val_addtototalprice - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "ليه رصيد"
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If

                    End If
                    If isdafedit = "notdefulit" Then
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                    'xtotalprice = val_storintotalprice - val_addtototalprice

                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    'rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                    'rs_Vendors.Update()
                    'End If
                End If

                If val_newtotalprice = val_Oldtotlprice Then
                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    rs_Vendors("Cust_Blance_Medin").Value = val_storintotalprice
                    rs_Vendors.Update()
                    'End If
                End If
            End If
            If cv = "vend" Then
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                    isdafedit = rs_Vendors("Vend_Blance_d").Value
                End If

                val_newtotalprice = toprice
                If val_Oldtotlprice < val_newtotalprice Then
                    val_addtototalprice = val_newtotalprice - val_Oldtotlprice
                    If isdafedit = "defulit" Then
                        If val_storintotalprice >= val_addtototalprice Then
                            xtotalprice = val_storintotalprice - val_addtototalprice
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice < val_addtototalprice Then
                            xtotalprice = val_addtototalprice - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "عليه رصيد"
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                    If isdafedit = "notdefulit" Then
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                End If

                If val_newtotalprice < val_Oldtotlprice Then
                    val_addtototalprice = val_Oldtotlprice - val_newtotalprice

                    If isdafedit = "defulit" Then
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()

                    End If
                    If isdafedit = "notdefulit" Then
                        If val_addtototalprice < val_storintotalprice Then
                            xtotalprice = val_storintotalprice - val_addtototalprice
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_addtototalprice >= val_storintotalprice Then
                            xtotalprice = val_addtototalprice - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "defulit"
                            qw = "ليه رصيد"
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If

                End If
                If val_newtotalprice = val_Oldtotlprice Then
                    rs_Vendors("Vend_Blance_Dein").Value = val_storintotalprice
                    rs_Vendors.Update()
                    'End If
                End If
            End If

            '' '' '' ''For payment and price in the export  and  blance tables
            ss1 = False
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            val_Blance_Total = rs_Store("Blance_Total").Value

            val_newpayment = payme

            If val_Oldpayment > val_newpayment Then
                val_addtopayment = val_Oldpayment - val_newpayment
                xpayment = val_Blance_Total - val_addtopayment
                If cv = "cust" Then
                    ss1 = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                        isdafedit = rs_Vendors("Cust_Blance_d").Value
                    End If

                    If isdafedit = "defulit" Then
                        xtotalprice = val_storintotalprice + val_addtopayment
                        rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                    If isdafedit = "notdefulit" Then

                        If val_storintotalprice > val_addtopayment Then
                            xtotalprice = val_storintotalprice - val_addtopayment
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice <= val_addtopayment Then
                            xtotalprice = val_addtopayment - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "defulit"
                            qw = "عليه رصيد"
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If

                    End If
                End If
                If cv = "vend" Then
                    ss1 = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                        isdafedit = rs_Vendors("Vend_Blance_d").Value
                    End If

                    If isdafedit = "defulit" Then
                        If val_storintotalprice >= val_addtopayment Then
                            xtotalprice = val_storintotalprice - val_addtopayment
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice < val_addtopayment Then
                            xtotalprice = val_addtopayment - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "عليه رصيد"
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                    If isdafedit = "notdefulit" Then
                        xtotalprice = val_storintotalprice + val_addtopayment
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                End If


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    rr = 0
                Else
                    rr = rs_StoreVend("Money_ID").Value
                    rs_StoreVend("Money_Price").Value = val_newpayment
                    rs_StoreVend.Update()
                End If

                Dim mm1 As String
                mm1 = "تعديل وارد"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = rr
                rs_Store("Blance_Total").Value = xpayment
                rs_Store("DelFlag").Value = ss1
                rs_Store("Blance_State").Value = mm1
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()
            End If

            ss1 = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                isdafedit = rs_Vendors("Cust_Blance_d").Value
            End If
            If val_newpayment > val_Oldpayment Then
                val_addtopayment = val_newpayment - val_Oldpayment
                xpayment = val_Blance_Total + val_addtopayment
                If cv = "cust" Then
                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    '    Dim jj As Integer
                    '    jj = rs_Vendors("Cust_Blance_Medin").Value - val_addtopayment
                    '    rs_Vendors("Cust_Blance_Medin").Value = jj
                    '    rs_Vendors.Update()
                    'End If
                    ss1 = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                        isdafedit = rs_Vendors("Cust_Blance_d").Value
                    End If

                    If isdafedit = "defulit" Then
                        If val_storintotalprice >= val_addtopayment Then
                            xtotalprice = val_storintotalprice - val_addtopayment
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice < val_addtopayment Then
                            xtotalprice = val_addtopayment - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "ليه رصيد"
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                    If isdafedit = "notdefulit" Then
                        xtotalprice = val_storintotalprice + val_addtopayment
                        rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                        rs_Vendors.Update()

                    End If
                End If
                If cv = "vend" Then
                    ss1 = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                        isdafedit = rs_Vendors("Vend_Blance_d").Value
                    End If

                    If isdafedit = "defulit" Then
                        xtotalprice = val_storintotalprice + val_addtopayment
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                    If isdafedit = "notdefulit" Then
                        If val_storintotalprice > val_addtopayment Then
                            xtotalprice = val_storintotalprice - val_addtopayment
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                        If val_storintotalprice <= val_addtopayment Then
                            xtotalprice = val_addtopayment - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "defulit"
                            qw = "ليه رصيد"
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If

                    End If
                End If


                If val_Oldpayment = z And val_newpayment > z Then
                    dy = DateTimePicker1.Value.Year
                    dm = DateTimePicker1.Value.Month
                    dd = DateTimePicker1.Value.Day
                    dstri = dy + "-" + dm + "-" + dd
                    ss1 = False
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                        If rs_Customers.State = 1 Then rs_Customers.Close()
                        rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        rs_Customers.AddNew()
                        'rs_Customers("Money_ID").Value = u_name1
                        rs_Customers("Money_Type").Value = money_Typeedit
                        rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
                        rs_Customers("Money_Datestri").Value = dstri
                        rs_Customers("Money_Price").Value = val_newpayment

                        If cv = "cust" Then
                            stredit = "وارد من " + ComboBox2.Text
                            strnoteedit = "رقم أذن الصادر للمخزن " + f1
                            rs_Customers("Cust_ID").Value = tt
                            rs_Customers("Vend_ID").Value = oe

                        End If
                        'If cv = "vend" Then
                        '    stredit = "وارد من " + ComboBox8.Text
                        '    strnoteedit = "رقم أذن الصادر للمخزن " + f1
                        '    rs_Customers("Cust_ID").Value = oe
                        '    rs_Customers("Vend_ID").Value = tt
                        'End If
                        rs_Customers("Money_Type_Arabic").Value = "وارد إلى الخزنة"
                        rs_Customers("Money_Reason").Value = stredit
                        rs_Customers("Money_Note").Value = strnoteedit
                        rs_Customers("User_ID").Value = u_Id
                        rs_Customers("Money_DelFlag").Value = ss1
                        rs_Customers("Month_ID").Value = month_id
                        rs_Customers("Stor_ID").Value = f
                        rs_Customers.Update()
                    Else
                        rs_StoreVend("Money_Price").Value = val_newpayment
                        rs_StoreVend.Update()
                    End If

                End If
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    rr = 0
                Else
                    rr = rs_StoreVend("Money_ID").Value
                    rs_StoreVend("Money_Price").Value = val_newpayment
                    rs_StoreVend.Update()
                End If
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                'Else
                '    rr = rs_StoreVend("Money_ID").Value
                '    rs_StoreVend("Money_Price").Value = val_newpayment
                '    rs_StoreVend.Update()
                'End If

                Dim mm1 As String
                mm1 = "تعديل وارد"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = rr
                rs_Store("Blance_Total").Value = xpayment
                rs_Store("DelFlag").Value = ss1
                rs_Store("Blance_State").Value = mm1
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()
            End If




            ' '' '' '' '' '' '' '' '' '' '' '' '' ''Update store vend
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            ss1 = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Stor Where Stor_ID=" & f & " And Stor_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'rs_StoreVend("Stor_From").Value = gov_Name
            'rs_StoreVend("Stor_To").Value = ComboBox6.Text
            'rs_StoreVend("car_Id").Value = car_id
            'rs_StoreVend("Drivers_Name").Value = Dri_Name
            rs_StoreVend("User_ID_edit").Value = u_Id
            'rs_StoreVend("Cust_ID").Value = cust_id
            rs_StoreVend("Stor_Date").Value = DateTimePicker1.Value.ToShortDateString()
            rs_StoreVend("Stor_Datestri").Value = dstri
            rs_StoreVend("Stor_Quntity_No").Value = quentity
            rs_StoreVend("Stor_Note").Value = TextBox1.Text
            rs_StoreVend("Stor_Type").Value = order_Typeedit
            rs_StoreVend("Month_ID").Value = month_id
            rs_StoreVend("Stor_Price_tin").Value = price
            rs_StoreVend("Stor_Total_Price").Value = toprice
            rs_StoreVend("Stor_Payment").Value = payme
            rs_StoreVend.Update()
            MsgBox("تم تعديل عملية الصادر بنجاح", MsgBoxStyle.Information, "تعديل بيانات الصادرات")
            ComboBox2.Text = ""

            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"

            ' '' '' '' '' '' ''Update store quntity

            ss1 = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                Dim mm55 As String
                mm55 = "تعديل صادر"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()

                rs_Store("Store_Quntity").Value = x
                rs_Store("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store("Stor_ID").Value = f
                rs_Store("Store_DelFlag").Value = ss1
                rs_Store("User_ID").Value = u_Id
                rs_Store("Stroe_State").Value = mm55
                rs_Store("Item_ID").Value = ite_id
                rs_Store.Update()
            Else
                Dim mm66 As String
                mm66 = "تعديل صادر"
                rs_StoreVend("Store_Quntity").Value = x
                rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_StoreVend("Stor_ID").Value = f
                rs_StoreVend("Store_DelFlag").Value = ss1
                rs_StoreVend("User_ID").Value = u_Id
                rs_StoreVend("Stroe_State").Value = mm66
                rs_StoreVend("Item_ID").Value = ite_id
                rs_StoreVend.Update()
            End If

            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Label27.Text = rs_Store("Blance_Total").Value
            End If

            Label1.Visible = True
            ComboBox2.Visible = True
            rs_Store.Close()
            'rs_Vendors.Close()
            'rs_StoreVend.Close()
            'rs_Customers.Close()
            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"
            'ComboBox2.Enabled = True
            ComboBox7.Enabled = True
            st_request = "notload"
            getallfetora()


            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox7.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""

            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"

            btn_etite.Enabled = False
            btn_delete.Enabled = False
            btn_add.Enabled = False
            Label7.Enabled = False
            'TextBox3.Enabled = False
            'Button1.Enabled = False
            btn_save.Enabled = True

            ComboBox7.Enabled = True
            Button2.Visible = True
            ComboBox7.Select()
            Dim s As Boolean
            s = False
            Try

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Count(Stor_ID) as df From Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                If rs_StoreVend("df").Value = 0 Then
                    Label11.Text = 1
                Else
                    Dim u_name As Integer
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    u_name = rs_StoreVend("Stor_ID").Value
                    Label11.Text = u_name + 1
                End If


                Dim s1 As Boolean
                s1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & Fetora_custid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox2.Text = rs_Vendors("Cust_Name").Value
                Label19.Text = rs_Vendors("Cust_Blance_Type").Value
                'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
                Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
                cust_id = Fetora_custid

                If rs_Cars.State = 1 Then rs_Cars.Close()
                rs_Cars.Open("Select * From year_monthes Where Month_ID=" & Fetora_monthid & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                ComboBox1.Text = rs_Cars("Month_Name").Value
                month_id = Fetora_monthid

                If fetore_type = "serch" Then


                    Dim order_Type As String
                    order_Type = "Cust"
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Fetora where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                    End If
                    DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                    'rs_StoreVend.Close()
                Else
                    DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            End Try


            rs_StoreVend.Close()
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If Label11.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        Dim st25 As String
        Try
            Dim s As Boolean
            s = False
            Dim tt As Integer
            If cv = "cust" Then
                tt = cust_iddelete
            End If
            If cv = "vend" Then
                tt = vend_iddelete
            End If

            Dim val_storintotalprice As Integer
            Dim val_Blance_Total2 As Integer
            Dim z As Integer
            z = 0

            Dim g As String
            g = MsgBox("هل تريد حذف هذا الاذن ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

            If g = vbYes Then

                Dim g1 As String
                g1 = MsgBox("هل تريد حذف المبلغ المدفوع أيضا ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g1 = vbYes Then
                    st25 = "yesDeletepayment"
                    If rs_Store.State = 1 Then rs_Store.Close()
                    rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Store.EOF Or rs_Store.BOF Then

                    Else
                        val_Blance_Total2 = rs_Store("Blance_Total").Value
                    End If

                    If cv = "cust" Then
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                            isdafdelete = rs_Vendors("Cust_Blance_d").Value
                        End If
                        'Dim ll As Integer
                        'Dim ll77 As Integer
                        'll77 = val_storintotalprice + val_Oldpayment
                        'll = ll77 - val_Oldtotlprice
                        'If ll < z Then
                        '    MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                        '    txt_quntity.Select()
                        '    Exit Sub
                        'End If

                        Dim ty433 As Integer
                        ty433 = val_Blance_Total2 - val_Oldpayment
                        If ty433 < z Then
                            MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                            txt_Payment.Select()
                            Exit Sub
                        End If
                    End If
                    If cv = "vend" Then
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                            isdafdelete = rs_Vendors("Vend_Blance_d").Value
                        End If
                    End If
                    Dim ty43 As Integer
                    ty43 = val_Blance_Total2 - val_Oldpayment
                    If ty43 < z Then
                        MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                        txt_Payment.Select()
                        Exit Sub
                    End If
                Else
                    st25 = "noDeletepayment"
                    If cv = "cust" Then
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                            isdafdelete = rs_Vendors("Cust_Blance_d").Value
                        End If
                        'Dim ll32 As Integer
                        'll32 = val_storintotalprice - val_Oldtotlprice
                        'If ll32 < z Then
                        '    MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                        '    txt_quntity.Select()
                        '    Exit Sub
                        'End If
                    End If
                    If cv = "vend" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                            isdafdelete = rs_Vendors("Vend_Blance_d").Value
                        End If
                    End If

                End If

                ' '' '' ''To sure about totalprice and payment not make blance of vendor < 0


                'Dim ll2 As Integer
                'll2 = val_storintotalprice - val_Oldtotlprice
                'Dim ty As Integer
                'ty = ll2 - val_Oldpayment
                'If ty < z Then
                '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '    txt_quntity.Select()
                '    Exit Sub
                'End If



                '' '' '' '' '' '' '' '' '' '' '' '' '' ''For  delete

                Dim r As Integer
                r = Label11.Text
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Stor Where Stor_ID=" & r & " And Stor_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim ne22 As Integer
                ne22 = rs_StoreVend("Stor_Quntity_No").Value

                If rs_Store.State = 1 Then rs_Store.Close()
                Dim s3 As Boolean
                s3 = False
                rs_Store.Open("Select * From Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim val_in As Integer
                val_in = rs_Store("Store_Quntity").Value

                'If val_in >= ne22 Then

                s = True
                rs_StoreVend("Stor_DelFlag").Value = s
                rs_StoreVend("User_ID_delete").Value = u_Id
                rs_StoreVend.Update()
                MsgBox("تم حذف الحركة بنجاح", MsgBoxStyle.Information, "حذف بيانات ")
                ComboBox2.Text = ""
                ComboBox1.Text = ""

                txt_quntity.Text = 0
                txt_price.Text = 0
                txt_Payment.Text = 0
                txt_totalPrice.Text = 0
                TextBox1.Text = ""
                Label11.Text = "0"
                ComboBox7.Text = ""

                Dim fin As Integer
                fin = val_in + ne22
                s3 = False
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    Dim mm As String
                    mm = "حذف صادر"
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_StoreVend.AddNew()

                    rs_StoreVend("Store_Quntity").Value = fin
                    rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                    rs_StoreVend("Stor_ID").Value = r
                    rs_StoreVend("Store_DelFlag").Value = s3
                    rs_StoreVend("User_ID").Value = u_Id
                    rs_StoreVend("Stroe_State").Value = mm
                    rs_StoreVend("Item_ID").Value = ite_id
                    rs_StoreVend.Update()
                Else
                    Dim mm As String
                    mm = "حذف صادر"

                    rs_StoreVend("Store_Quntity").Value = fin
                    rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                    rs_StoreVend("Stor_ID").Value = r
                    rs_StoreVend("Store_DelFlag").Value = s3
                    rs_StoreVend("User_ID").Value = u_Id
                    rs_StoreVend("Stroe_State").Value = mm
                    rs_StoreVend("Item_ID").Value = ite_id
                    rs_StoreVend.Update()
                End If

                If st25 = "yesDeletepayment" Then
                    Dim rr As Integer
                    s3 = False
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & r & " And Money_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                    Else
                        s = True
                        rr = rs_StoreVend("Money_ID").Value
                        rs_StoreVend("Money_DelFlag").Value = s
                        rs_StoreVend.Update()
                    End If
                    s3 = False
                    Dim fin2 As Integer
                    Dim val_Blance_Total As Integer
                    If rs_Store.State = 1 Then rs_Store.Close()
                    rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s3 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    val_Blance_Total = rs_Store("Blance_Total").Value
                    fin2 = val_Blance_Total - val_Oldpayment
                    Dim mm1 As String
                    mm1 = "حذف وارد"
                    If rs_Store.State = 1 Then rs_Store.Close()
                    rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_Store.AddNew()
                    'rs_Store("ID").Value = u_name21
                    rs_Store("Money_ID").Value = rr
                    rs_Store("Blance_Total").Value = fin2
                    rs_Store("DelFlag").Value = s3
                    rs_Store("Blance_State").Value = mm1
                    rs_Store("User_ID").Value = u_Id
                    rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                    rs_Store.Update()

                    If cv = "cust" Then
                        ' ''عليه رصيد
                        If isdafdelete = "defulit" Then
                            s3 = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim jj As Integer
                                jj = val_storintotalprice + val_Oldpayment
                                rs_Vendors("Cust_Blance_Medin").Value = jj
                                rs_Vendors.Update()
                            End If



                        End If

                        ' ''ليه رصيد
                        If isdafdelete = "notdefulit" Then
                            If val_storintotalprice > val_Oldpayment Then
                                Dim ff As Integer
                                ff = val_storintotalprice - val_Oldpayment
                                s3 = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    'Dim jj As Integer
                                    'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                    rs_Vendors("Cust_Blance_Medin").Value = ff
                                    rs_Vendors.Update()
                                End If
                            End If

                            If val_storintotalprice <= val_Oldpayment Then
                                Dim ff8 As Integer
                                ff8 = val_Oldpayment - val_storintotalprice
                                s3 = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim re As String
                                    Dim qw As String
                                    re = "defulit"
                                    qw = "عليه رصيد"
                                    'Dim jj As Integer
                                    'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                    rs_Vendors("Cust_Blance_Medin").Value = ff8
                                    rs_Vendors("Cust_Blance_Type").Value = qw
                                    rs_Vendors("Cust_Blance_d").Value = re
                                    rs_Vendors.Update()
                                End If
                            End If
                        End If
                    End If
                    If cv = "vend" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                            isdafdelete = rs_Vendors("Vend_Blance_d").Value
                        End If
                        ' ''ليه رصيد
                        If isdafdelete = "defulit" Then


                            If val_storintotalprice >= val_Oldpayment Then
                                s = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim jj As Integer
                                    jj = val_storintotalprice - val_Oldpayment
                                    rs_Vendors("Vend_Blance_Dein").Value = jj
                                    rs_Vendors.Update()
                                End If
                            End If
                            If val_storintotalprice < val_Oldpayment Then
                                s = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim jj As Integer
                                    jj = val_Oldpayment - val_storintotalprice
                                    Dim re As String
                                    Dim qw As String
                                    re = "notdefulit"
                                    qw = "عليه رصيد"
                                    rs_Vendors("Vend_Blance_Dein").Value = jj
                                    rs_Vendors("Vend_Blance_Type").Value = qw
                                    rs_Vendors("Vend_Blance_d").Value = re
                                    rs_Vendors.Update()
                                End If
                            End If
                        End If
                        ' ''عليه رصيد
                        If isdafdelete = "notdefulit" Then

                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim jj As Integer
                                jj = val_storintotalprice + val_Oldpayment
                                rs_Vendors("Vend_Blance_Dein").Value = jj
                                rs_Vendors.Update()
                            End If
                        End If
                    End If

                End If

                If st25 = "noDeletepayment" Then
                    s3 = False
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & r & " And Money_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                    Else
                        Dim tu As Integer
                        tu = 0
                        rs_StoreVend("Stor_ID").Value = tu
                        rs_StoreVend.Update()
                    End If
                End If

                If cv = "cust" Then
                    s = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                        isdafdelete = rs_Vendors("Cust_Blance_d").Value
                    End If
                    s3 = False
                    ' ''عليه رصيد
                    If isdafdelete = "defulit" Then
                        If val_storintotalprice >= val_Oldtotlprice Then
                            Dim ff As Integer
                            ff = val_storintotalprice - val_Oldtotlprice
                            s3 = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                'Dim jj As Integer
                                'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                rs_Vendors("Cust_Blance_Medin").Value = ff
                                rs_Vendors.Update()
                            End If
                        End If

                        If val_storintotalprice < val_Oldtotlprice Then
                            Dim ff8 As Integer
                            ff8 = val_Oldtotlprice - val_storintotalprice
                            s3 = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim re As String
                                Dim qw As String
                                re = "notdefulit"
                                qw = "ليه رصيد"
                                rs_Vendors("Cust_Blance_Medin").Value = ff8
                                rs_Vendors("Cust_Blance_Type").Value = qw
                                rs_Vendors("Cust_Blance_d").Value = re
                                rs_Vendors.Update()
                            End If
                        End If


                    End If

                    ' ''ليه رصيد
                    If isdafdelete = "notdefulit" Then
                        s3 = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            Dim jj As Integer
                            jj = val_storintotalprice + val_Oldtotlprice
                            rs_Vendors("Cust_Blance_Medin").Value = jj
                            rs_Vendors.Update()
                        End If

                    End If
                End If
                If cv = "vend" Then
                    s = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                        isdafdelete = rs_Vendors("Vend_Blance_d").Value
                    End If
                    ' ''ليه رصيد
                    If isdafdelete = "defulit" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            Dim jj As Integer
                            jj = val_storintotalprice + val_Oldtotlprice
                            rs_Vendors("Vend_Blance_Dein").Value = jj
                            rs_Vendors.Update()
                        End If

                    End If
                    ' ''عليه رصيد
                    If isdafdelete = "notdefulit" Then
                        If val_storintotalprice > val_Oldtotlprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim jj As Integer
                                jj = val_storintotalprice - val_Oldtotlprice
                                rs_Vendors("Vend_Blance_Dein").Value = jj
                                rs_Vendors.Update()
                            End If
                        End If
                        If val_storintotalprice <= val_Oldtotlprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim jj As Integer
                                jj = val_Oldtotlprice - val_storintotalprice
                                Dim re As String
                                Dim qw As String
                                re = "defulit"
                                qw = "ليه رصيد"
                                rs_Vendors("Vend_Blance_Dein").Value = jj
                                rs_Vendors("Vend_Blance_Type").Value = qw
                                rs_Vendors("Vend_Blance_d").Value = re
                                rs_Vendors.Update()


                            End If
                        End If
                    End If
                End If

                'End If
                'If val_in < ne22 Then
                '    MsgBox("رصيد المخزن لا يسمح بارجاع شحنة الوارد")
                'End If




                Dim sMoney_Blance As Boolean
                sMoney_Blance = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    If rs_Store.State = 1 Then rs_Store.Close()
                    rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    Label27.Text = rs_Store("Blance_Total").Value
                End If

                Label1.Visible = True
                ComboBox2.Visible = True
                rs_Store.Close()
                'rs_StoreVend.Close()
                'rs_Vendors.Close()
                Label15.Text = "-"
                Label16.Text = "-"
                Label11.Text = "-"
                'ComboBox2.Enabled = True
                ComboBox7.Enabled = True
                st_request = "notload"
                getallfetora()

                Label1.Visible = True
                ComboBox2.Visible = True
                ComboBox7.Text = ""
                ComboBox1.Text = ""
                ComboBox2.Text = ""

                txt_quntity.Text = 0
                txt_quntity.Text = 0
                txt_Payment.Text = 0
                txt_price.Text = 0
                TextBox1.Text = ""
                Label15.Text = "-"
                Label16.Text = "-"

                btn_etite.Enabled = False
                btn_delete.Enabled = False
                btn_add.Enabled = False
                Label7.Enabled = False
                'TextBox3.Enabled = False
                'Button1.Enabled = False
                btn_save.Enabled = True

                ComboBox7.Enabled = True
                Button2.Visible = True
                ComboBox7.Select()
                'Dim s As Boolean
                s = False
                Try

                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select Count(Stor_ID) as df From Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                    If rs_StoreVend("df").Value = 0 Then
                        Label11.Text = 1
                    Else
                        Dim u_name As Integer
                        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                        rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        u_name = rs_StoreVend("Stor_ID").Value
                        Label11.Text = u_name + 1
                    End If


                    Dim s1 As Boolean
                    s1 = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & Fetora_custid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    ComboBox2.Text = rs_Vendors("Cust_Name").Value
                    Label19.Text = rs_Vendors("Cust_Blance_Type").Value
                    'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
                    Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
                    cust_id = Fetora_custid

                    If rs_Cars.State = 1 Then rs_Cars.Close()
                    rs_Cars.Open("Select * From year_monthes Where Month_ID=" & Fetora_monthid & " And Month_DeleFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    ComboBox1.Text = rs_Cars("Month_Name").Value
                    month_id = Fetora_monthid

                    If fetore_type = "serch" Then


                        Dim order_Type As String
                        order_Type = "Cust"
                        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                        rs_StoreVend.Open("Select * From Fetora where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                        If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                        End If
                        DateTimePicker1.Value = rs_StoreVend("Fetora_date").Value
                        'rs_StoreVend.Close()
                    Else
                        DateTimePicker1.Value = Frm_StoreCust.DateTimePicker1.Value
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
                End Try


                rs_StoreVend.Close()
                rs_Vendors.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub





    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox1.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox1.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub






    Private Sub txt_quntity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_quntity.TextChanged
        'Dim t As Integer
        't = 1000
        'If Val(txt_quntity.Text) > t Then
        '    Label12.Text = "طـن"
        'ElseIf Val(txt_quntity.Text) <= t Then
        '    Label12.Text = "كيلـو"
        'End If
        Try
            If Label30.Text = "وحدة" Then
                Dim tt5 As Integer
                tt5 = Val(txt_price.Text) * Val(txt_quntity.Text)
                txt_totalPrice.Text = tt5
            End If
            If Label30.Text = "طن" Then
                Dim dd1 As Double
                dd1 = Val(txt_price.Text) / 1000
                Dim tt22 As Integer
                tt22 = dd1 * Val(txt_quntity.Text)
                txt_totalPrice.Text = tt22
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
        'If button_Events = "new" Then
        '    Label15.Text = Val(Label15.Text) - Val(txt_quntity.Text)
        '    Label16.Text = Val(Label16.Text) + Val(tt)
        'End If
        'Dim gf As Integer
        'Dim adto As Integer
        'gf = Val(txt_quntity.Text)
        'If button_Events = "edit" Then

        '    If val_Old > gf Then
        '        adto = val_Old - gf
        '        Label15.Text = store_quntity + adto
        '    End If
        '    If val_Old < gf Then
        '        adto = gf - val_Old
        '        Label15.Text = store_quntity - adto
        '    End If
        'End If

    End Sub

    Private Sub txt_price_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_price.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_price_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_price.TextChanged
        Try
            If Label30.Text = "وحدة" Then
                Dim tt6 As Integer
                tt6 = Val(txt_price.Text) * Val(txt_quntity.Text)
                txt_totalPrice.Text = tt6
            End If
            If Label30.Text = "طن" Then
                Dim dd1 As Double
                dd1 = Val(txt_price.Text) / 1000
                Dim tt22 As Integer
                tt22 = dd1 * Val(txt_quntity.Text)
                txt_totalPrice.Text = tt22
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox7_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Itemes Where Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox7.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox7.Items.Add(rs_Vendors("Item_Name").Value)
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
            rs_Vendors.Open("Select * From Itemes Where Item_Name='" & u_name & "' And Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ite_id = rs_Vendors("Item_ID").Value
            Dim jh As String
            jh = rs_Vendors("Item_Wight_Type").Value
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                Dim mm As String
                mm = "أول مرة"
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Vendors.AddNew()

                rs_Vendors("Store_Quntity").Value = 0
                rs_Vendors("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Vendors("Stor_ID").Value = 0
                rs_Vendors("Store_DelFlag").Value = s
                rs_Vendors("User_ID").Value = u_Id
                rs_Vendors("Stroe_State").Value = mm
                rs_Vendors("Item_ID").Value = ite_id
                rs_Vendors("Item_Wight_Type").Value = jh
                rs_Vendors.Update()
                Label15.Text = 0
                'store_quntity = 0
                If jh = "وحدة" Then
                    Label30.Text = jh
                    Label18.Text = jh
                    Label21.Text = "سعر الوحدة :"
                End If
                If jh = "كيلو" Then
                    Label30.Text = "طن"
                    Label18.Text = "كيلو"
                    Label21.Text = "سعر الطن :"
                End If
            Else

                Label15.Text = rs_Vendors("Store_Quntity").Value
                jh = rs_Vendors("Item_Wight_Type").Value
                If jh = "وحدة" Then
                    Label30.Text = jh
                    Label18.Text = jh
                    Label21.Text = "سعر الوحدة :"
                End If
                If jh = "كيلو" Then
                    Label30.Text = "طن"
                    Label18.Text = "كيلو"
                    Label21.Text = "سعر الطن :"
                End If
                'store_quntity = rs_Vendors("Store_Quntity").Value
            End If
            txt_price.Text = ""
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_Payment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Payment.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub



   
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Cursor = Cursors.WaitCursor
        Dim rpt As New Rpt_Show_fetora
        Dim frm_rpt As New frm_rpt_showfetora
        rdat.Clear()
        'SQLEXPRESS
        rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
        rcon.Open()
        Dim ssd As Boolean
        ssd = False
        Dim tp As String
        tp = "Cust"
        Dim T_fato As String
        Dim V_fato As String
        Dim C_Name As String
        rcomand = New SqlCommand("Select sm.Stor_ID, (Select vend.Cust_Name From Customers as vend Where vend.Cust_ID =" & cust_id & ") as custname,(Select vend.Item_Name From Itemes as vend Where vend.Item_ID = sm.Item_ID) as itemname, sm.Stor_Date, sm.Stor_Quntity_No, sm.Stor_Price_tin, sm.Stor_Total_Price, sm.Stor_Payment, sm.Stor_Note From Stor as sm Where sm.Cust_ID =" & cust_id & " And sm.Fetora_id=" & Fetora_id & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "' ORDER BY sm.Stor_Date ASC", rcon)
        'rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select Sum(sm.Stor_Total_Price) as c_No From Stor as sm Where sm.Cust_ID =" & cust_id & " And sm.Fetora_id=" & Fetora_id & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If IsDBNull(rs_StoreVend("c_No").Value) Then
            T_fato = 0
        Else
            T_fato = rs_StoreVend("c_No").Value
        End If

        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select Sum(sm.Stor_Payment) as c_No FROM Stor as sm Where sm.Cust_ID =" & cust_id & " And sm.Fetora_id=" & Fetora_id & " And sm.Stor_Type='" & tp & "' And sm.Stor_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If IsDBNull(rs_StoreVend("c_No").Value) Then
            V_fato = 0
        Else
            V_fato = rs_StoreVend("c_No").Value
        End If
        Dim s As Boolean
        s = False
        Dim ghf As Integer
        Dim sy As String
        sy = "الصادرات"
        Dim sy3 As String
        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        ghf = rs_StoreVend("Cust_Blance_Medin").Value
        sy3 = rs_StoreVend("Cust_Blance_Type").Value
        C_Name = rs_StoreVend("Cust_Name").Value
        rs_StoreVend.Close()
        Dim fg As String

        Dim gh As Integer

        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select * From Fetora Where Fetora_id=" & Fetora_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        gh = rs_StoreVend("Fetora_blancebefore").Value
        fg = rs_StoreVend("Fetora_blancebeforetype").Value
        rada = New SqlDataAdapter(rcomand)

        rada.Fill(rdat, "Stor")
        rpt.SetDataSource(rdat)
        rpt.SetParameterValue("MonthNameun", user_name)
        rpt.SetParameterValue("T_Name2", sy3)
        rpt.SetParameterValue("B_name", ghf)
        rpt.SetParameterValue("T_namem", V_fato)
        rpt.SetParameterValue("T_nameo", gh)
        rpt.SetParameterValue("T_namet", fg)
        rpt.SetParameterValue("C_name", C_Name)
        rpt.SetParameterValue("T_Name1", T_fato)
        frm_rpt.CrystalReportViewer1.ReportSource = rpt
        frm_rpt.CrystalReportViewer1.Refresh()
        frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
        Dim frm As New Form
        With frm
            .Controls.Add(frm_rpt.CrystalReportViewer1)
            .Text = "طباعة فاتورة"
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Customers Where Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox2.Items.Add(rs_Vendors("Cust_Name").Value)
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
            rs_Vendors.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            cust_id = rs_Vendors("Cust_ID").Value
            Label19.Text = rs_Vendors("Cust_Blance_Type").Value
            cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            'If rs_Vendors.State = 1 Then rs_Vendors.Close()
            'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value

            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    
End Class