Imports System.Data.OleDb
Imports ADODB
Imports System.Data.SqlClient
Imports System.Data
Public Class frm_StoreCust_Cdid
    Public cust_state_defult As String
    Public oe As Integer
    Public ite_id As Integer
    Public month_id As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    Public isdafsave As String
    Private Sub frm_StoreCust_Cdid_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
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
            auth_add = rs_auth("User_storvend_add").Value
            If auth_add = ut Then
                btn_add.Visible = False
                btn_save.Visible = False
            Else
                btn_add.Visible = True
                btn_save.Visible = True
            End If
            If u_Id = 1 Then
                btn_add.Visible = True
                btn_save.Visible = True
                Label17.Visible = True
                Label15.Visible = True
                Label18.Visible = True
                Label29.Visible = True
                Label27.Visible = True
                Label28.Visible = True
            End If
            Dim s1 As Boolean
            s1 = False
            oe = 1
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox8.Text = rs_Vendors("Cust_Name").Value
            Label19.Text = rs_Vendors("Cust_Blance_Type").Value
            'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value

            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Cust"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & s & "' And Stor_ID=" & sid_Cust_Cdid & " And Stor_Type='" & order_Type & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                Exit Sub
            End If
            DateTimePicker1.Value = rs_StoreVend("Stor_Date").Value

            Dim s15 As Boolean
            s15 = False
            Dim gh5 As Integer
            gh5 = rs_StoreVend("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.Open("Select * From year_monthes Where Month_ID=" & gh5 & " And Month_DeleFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox1.Text = rs_U("Month_Name").Value
            month_id = gh5

            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                Exit Sub
            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Label27.Text = rs_Store("Blance_Total").Value
            End If
            getvendname()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click

        Try
            ComboBox7.Text = ""
            ComboBox1.Text = ""

            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"

            btn_add.Enabled = False
            btn_save.Enabled = True

            Dim s1 As Boolean
            s1 = False

            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox8.Text = rs_Vendors("Cust_Name").Value
            Label19.Text = rs_Vendors("Cust_Blance_Type").Value
            'cust_state_defult = rs_Vendors("Cust_Blance_d").Value
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value

            ComboBox7.Select()

            Dim s As Boolean
            s = False
            Dim order_Type As String
            order_Type = "Cust"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & s & "' And Stor_ID=" & sid_Cust_Cdid & " And Stor_Type='" & order_Type & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                Exit Sub
            End If
            DateTimePicker1.Value = rs_StoreVend("Stor_Date").Value

            Dim s15 As Boolean
            s15 = False
            Dim gh5 As Integer
            gh5 = rs_StoreVend("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.Open("Select * From year_monthes Where Month_ID=" & gh5 & " And Month_DeleFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox1.Text = rs_U("Month_Name").Value
            month_id = gh5



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
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
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

    Private Sub txt_quntity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_quntity.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_quntity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_quntity.TextChanged
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

    Private Sub txt_Payment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Payment.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_Payment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Payment.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            getvendname()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 160
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 100
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 120
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 120
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 330
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            Dim stt As String
            stt = "SELECT st.Stor_ID as [رقم الأذن], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Itemes as itm, year_monthes as mn where st.Cust_ID=" & custid_cdid & " And st.Stor_ID_Cdid=" & sid_Cust_Cdid & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And itm.Item_DelFlag='" & s & "' ORDER BY st.Stor_ID DESC"
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT SUM(st.Stor_Quntity_No) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    TextBox2.Text = 0
            'Else
            '    TextBox2.Text = rs_StoreVend("c_No").Value
            'End If
            'TextBox2.Text = rs_StoreVend("c_No").Value
            'rs_StoreVend.Close()

            'If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            'rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd, Itemes as itm, year_monthes as mn where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Item_ID = itm.Item_ID And st.Stor_DelFlag='" & s & "' And st.Stor_Type='" & order_Type & "' And mn.Month_ID=st.Month_ID And mn.Month_DeleFlag='" & s & "' And vd.Vend_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreCust("c_No").Value) Then
            '    Label3.Text = 0
            'Else
            '    Label3.Text = rs_StoreCust("c_No").Value
            'End If
            'Label3.Text = rs_StoreCust("c_No").Value
            'rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
       
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If ComboBox7.Text = "" Then
            MsgBox("من فضلك أختر أسم الصنف", MsgBoxStyle.Information, "تنبيه")
            ComboBox7.Select()
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
        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
        End If

        Try
            Dim s As Boolean
            s = False
            'Dim u_name As Integer
            'Dim u_name1 As Integer
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreVend.Open("Select * From Stor Where Stor_ID= (SELECT MAX(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'u_name = rs_StoreVend("Stor_ID").Value
            'u_name1 = u_name + 1
            Dim quentity As Integer
            quentity = Val(txt_quntity.Text)
            Dim toprice As Integer
            toprice = Val(txt_totalPrice.Text)
            Dim price As Integer
            price = Val(txt_price.Text)
            Dim payme As Integer
            payme = Val(txt_Payment.Text)
            Dim order_Type As String
            order_Type = "Vend"
            'Dim vendname As String
            'vendname = ComboBox2.Text
            Dim val_storintotalprice As Integer
            'Dim val_Blance_Total As Integer
            'Dim val_pay As Integer

            Dim z As Integer
            z = 0
            ' '' '' ''To sure about payment not make blance of Vendores < 0
            'Dim tt As Integer
            'tt = custid_cdid
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                isdafsave = rs_Vendors("Cust_Blance_d").Value
            End If
            'If rs_Store.State = 1 Then rs_Store.Close()
            'rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If rs_Store.EOF Or rs_Store.BOF Then

            'Else
            '    val_Blance_Total = rs_Store("Blance_Total").Value
            'End If
            '' ''عليه رصيد
            'If isdaf = "defulit" Then
            '    val_pay = txt_Payment.Text
            '    Dim ty2 As Integer
            '    Dim ty5 As Integer
            '    ty5 = val_storintotalprice + toprice
            '    ty2 = ty5 - txt_Payment.Text

            '    If ty2 < z Then
            '        MsgBox("رصيد المورد يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
            '        txt_quntity.Select()
            '        Exit Sub
            '    End If

            'End If

            '' ''ليه رصيد
            'If isdaf = "notdefulit" Then

            'End If
           
            'Dim mo As Integer
            'mo = val_Blance_Total - payme
            'If mo < z Then
            '    MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
            '    txt_Payment.Select()
            '    Exit Sub
            'End If
            '' '' '' ''--------------------------------------------
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim uk As Integer
            uk = 2
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
            rs_StoreVend("Vend_ID").Value = oe
            rs_StoreVend("Cust_ID").Value = custid_cdid
            rs_StoreVend("Stor_Date").Value = DateTimePicker1.Value.ToShortDateString()
            rs_StoreVend("Stor_Datestri").Value = dstri
            rs_StoreVend("Stor_Quntity_No").Value = quentity
            rs_StoreVend("Stor_DelFlag").Value = s
            rs_StoreVend("Item_ID").Value = ite_id
            rs_StoreVend("Month_ID").Value = month_id
            rs_StoreVend("Stor_Note").Value = TextBox1.Text
            rs_StoreVend("Stor_Type").Value = order_Type
            rs_StoreVend("Stor_Price_tin").Value = price
            rs_StoreVend("Stor_Total_Price").Value = toprice
            rs_StoreVend("Stor_ID_Cdid").Value = sid_Cust_Cdid
            rs_StoreVend("Stor_Payment").Value = z
            rs_StoreVend("Stor_Type_Arabic").Value = "وارد من"
            rs_StoreVend.Update()
            MsgBox("تم حفظ عملية الوارد بنجاح", MsgBoxStyle.Information, "حفظ بيانات الواردات")

            ComboBox1.Text = ""
            
            ComboBox7.Text = ""
            txt_quntity.Text = 0
            txt_price.Text = 0
            txt_Payment.Text = 0
            txt_totalPrice.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"
            btn_add.Enabled = True

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
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                ne = rs_Store("Store_Quntity").Value
                fin = ne + ne22
            End If


            Dim s8 As Boolean
            s8 = False
            Dim u_name8 As Integer
            Dim n_Type As String
            n_Type = "Vend"
            Dim mm As String
            mm = "حفظ وارد"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select* From Stor Where Stor_Type ='" & n_Type & "' And Stor_DelFlag='" & s8 & "' And Stor_ID = (SELECT MAx(Stor_ID)  FROM Stor)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            u_name8 = rs_StoreVend("Stor_ID").Value
            s = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then

                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_StoreVend.AddNew()

                rs_StoreVend("Store_Quntity").Value = fin
                rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_StoreVend("Stor_ID").Value = u_name8
                rs_StoreVend("Store_DelFlag").Value = s3
                rs_StoreVend("User_ID").Value = u_Id
                rs_StoreVend("Stroe_State").Value = mm
                rs_StoreVend("Item_ID").Value = ite_id
                rs_StoreVend.Update()
            Else

                rs_StoreVend("Store_Quntity").Value = fin
                rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_StoreVend("Stor_ID").Value = u_name8
                rs_StoreVend("Store_DelFlag").Value = s3
                rs_StoreVend("User_ID").Value = u_Id
                rs_StoreVend("Stroe_State").Value = mm
                rs_StoreVend("Item_ID").Value = ite_id
                rs_StoreVend.Update()
            End If

            ' ''عليه رصيد
            If isdafsave = "defulit" Then

                If val_storintotalprice >= toprice Then
                    Dim ff As Integer
                    ff = val_storintotalprice - toprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Cust_Blance_Medin").Value = ff
                        rs_Vendors.Update()
                    End If
                End If

                If val_storintotalprice < toprice Then
                    Dim ff8 As Integer
                    ff8 = toprice - val_storintotalprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
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
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & custid_cdid & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    Dim jj As Integer
                    jj = val_storintotalprice + toprice
                    rs_Vendors("Cust_Blance_Medin").Value = jj
                    rs_Vendors.Update()
                End If
            End If

           



            'If payme > 0 Then
            '    dy = DateTimePicker1.Value.Year
            '    dm = DateTimePicker1.Value.Month
            '    dd = DateTimePicker1.Value.Day
            '    dstri = dy + "-" + dm + "-" + dd
            '    order_Type = "Exp"
            '    Dim str As String
            '    str = "صادر إلى " + vendname
            '    Dim strnote As String
            '    Dim strn2 As String
            '    strn2 = u_name8
            '    strnote = "رقم أذن الوارد للمخزن " + strn2
            '    s = False
            '    If rs_Customers.State = 1 Then rs_Customers.Close()
            '    rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    rs_Customers.AddNew()
            '    'rs_Customers("Money_ID").Value = u_name1
            '    rs_Customers("Money_Type").Value = order_Type
            '    rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
            '    rs_Customers("Money_Datestri").Value = dstri
            '    rs_Customers("Money_Price").Value = payme
            '    rs_Customers("Money_Reason").Value = str
            '    rs_Customers("Money_Note").Value = strnote
            '    rs_Customers("User_ID").Value = u_Id
            '    rs_Customers("Money_DelFlag").Value = s
            '    rs_Customers("Month_ID").Value = month_id
            '    rs_Customers("Stor_ID").Value = u_name8
            '    rs_Customers("Vend_ID").Value = vend_id
            '    rs_Customers.Update()
            '    s3 = False
            '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
            '    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    If rs_Vendors.EOF Or rs_Vendors.BOF Then

            '    Else
            '        Dim jj As Integer
            '        jj = rs_Vendors("Vend_Blance_Dein").Value - payme
            '        rs_Vendors("Vend_Blance_Dein").Value = jj
            '        rs_Vendors.Update()
            '    End If
            '    s8 = False
            '    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            '    rs_StoreVend.Open("Select* From Imp_Exp_Money Where Money_Type ='" & order_Type & "' And Money_DelFlag='" & s8 & "' And Money_ID = (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    u_name8 = rs_StoreVend("Money_ID").Value

            '    'Dim s3 As Boolean
            '    's3 = False
            '    Dim ne9 As Integer
            '    Dim ne229 As Integer
            '    Dim fin9 As Integer
            '    s3 = False
            '    If rs_Store.State = 1 Then rs_Store.Close()
            '    rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    ne9 = rs_Store("Blance_Total").Value
            '    ne229 = rs_StoreVend("Money_Price").Value
            '    fin9 = ne9 - ne229

            '    Dim mm12 As String
            '    mm12 = "حفظ صادر"
            '    If rs_Store.State = 1 Then rs_Store.Close()
            '    rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    rs_Store.AddNew()
            '    'rs_Store("ID").Value = u_name21
            '    rs_Store("Money_ID").Value = u_name8
            '    rs_Store("Blance_Total").Value = fin9
            '    rs_Store("DelFlag").Value = s3
            '    rs_Store("Blance_State").Value = mm12
            '    rs_Store("User_ID").Value = u_Id
            '    rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
            '    rs_Store.Update()

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

            rs_Store.Close()
            rs_StoreVend.Close()
            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"

            getvendname()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub
End Class