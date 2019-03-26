Imports ADODB
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data
Public Class frm_custreview
    Public cust_id As Integer
    Public rrr As Integer = 0
    Public rrrss As String
    Public fd As Integer
    Public sw1 As String
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet

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

            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub frm_custreview_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If gcon.State = 1 Then gcon.Close()
            gcon.Open()
            'ComboBox1.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub dgvfatora()
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
    Private Sub dgvorder()
        Try
            DataGridView2.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(0).Width = 70
            DataGridView2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(1).Width = 70
            DataGridView2.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(2).Width = 90
            DataGridView2.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView2.Columns(3).Width = 100
            DataGridView2.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(4).Width = 100
            DataGridView2.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(5).Width = 100
            DataGridView2.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(6).Width = 100
            DataGridView2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(7).Width = 100
            DataGridView2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(8).Width = 100
            DataGridView2.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(9).Width = 100
            DataGridView2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView2.Columns(10).Width = 150
            DataGridView2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub
    Private Sub dgvmoney()
        Try
            DataGridView3.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(0).Width = 70
            DataGridView3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(1).Width = 110
            DataGridView3.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(2).Width = 110
            DataGridView3.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView3.Columns(3).Width = 100
            DataGridView3.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(4).Width = 170
            DataGridView3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(5).Width = 190
            DataGridView3.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub
    Private Sub get_fetora(ByVal sw As String)
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String


            stt = "SELECT fo.Fetora_id as [رقم الفاتورة], cu.Cust_Name as [أسم العميل], fo.Fetora_date as [تاريخ الفاتورة ], fo.Fetora_blancebeforetype as [نوع قبل], fo.Fetora_blancebefore as [رصيد قبل], fo.Fetora_totalprice as [إجمالى الفاتورة], fo.Fetora_totalpay as [إجمالى المدفوع], fo.Fetora_blanceaftertype as [نوع بعد], fo.Fetora_blanceafter as [رصيد بعد] FROM Fetora as fo,Customers as cu where cu.Cust_DelFlag='" & s & "' And cu.Cust_ID=" & cust_id & " And fo.Fetora_custid=" & cust_id & " And cu.Cust_ID=fo.Fetora_custid ORDER BY fo.Fetora_id " + sw



            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            dgvfatora()
            DataGridView1.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_order(ByVal sw As String)
        Try
            sw1 = sw
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String

            stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic as [نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm where st.Cust_ID=vd.Cust_ID And st.Item_ID=itm.Item_ID And st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_DelFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' ORDER BY st.Stor_Date " + sw

            'stt = "SELECT fo.Fetora_id as [رقم الفاتورة], cu.Cust_Name as [أسم العميل], fo.Fetora_date as [تاريخ الفاتورة ], fo.Fetora_blancebeforetype as [نوع قبل], fo.Fetora_blancebefore as [رصيد قبل], fo.Fetora_totalprice as [إجمالى الفاتورة], fo.Fetora_totalpay as [إجمالى المدفوع], fo.Fetora_blanceaftertype as [نوع بعد], fo.Fetora_blanceafter as [رصيد بعد] FROM Fetora as fo,Customers as cu where cu.Cust_DelFlag='" & s & "' And cu.Cust_ID=" & cust_id & " And fo.Fetora_custid=" & cust_id & " And cu.Cust_ID=fo.Fetora_custid ORDER BY fo.Fetora_id " + sw



            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView2.DataSource = dt.DefaultView
            gcon.Close()
            dgvorder()
            DataGridView2.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_order_fetoraid(ByVal sw As String, ByVal ghh As Integer)
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String

            stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm where st.Cust_ID=vd.Cust_ID And st.Item_ID=itm.Item_ID And st.Fetora_id=" & ghh & " And st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_DelFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' ORDER BY st.Stor_Date " + sw

            'stt = "SELECT fo.Fetora_id as [رقم الفاتورة], cu.Cust_Name as [أسم العميل], fo.Fetora_date as [تاريخ الفاتورة ], fo.Fetora_blancebeforetype as [نوع قبل], fo.Fetora_blancebefore as [رصيد قبل], fo.Fetora_totalprice as [إجمالى الفاتورة], fo.Fetora_totalpay as [إجمالى المدفوع], fo.Fetora_blanceaftertype as [نوع بعد], fo.Fetora_blanceafter as [رصيد بعد] FROM Fetora as fo,Customers as cu where cu.Cust_DelFlag='" & s & "' And cu.Cust_ID=" & cust_id & " And fo.Fetora_custid=" & cust_id & " And cu.Cust_ID=fo.Fetora_custid ORDER BY fo.Fetora_id " + sw



            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView2.DataSource = dt.DefaultView
            gcon.Close()
            dgvorder()
            DataGridView2.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_money(ByVal sw As String)
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim stt As String
            Dim ze As Integer
            ze = 0
            'stt = "SELECT st.Stor_ID as [رقم الأذن],st.Fetora_id as [رقم الفاتورة], st.Stor_Type_Arabic[نوع الأذن], vd.Cust_Name as [أسم العميل], itm.Item_Name as [أسم الصنف], st.Stor_Date as [تاريخ الأذن], st.Stor_Quntity_No as [الكمية], st.Stor_Price_tin as [السعر], st.Stor_Total_Price as [إجمالى السعر], st.Stor_Payment as [المبلغ المدفوع], st.Stor_Note as [ملاحظات] FROM Stor as st, Customers as vd, Itemes as itm where st.Cust_ID=vd.Cust_ID And st.Item_ID=itm.Item_ID And st.Fetora_id=" & ghh & " And st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_DelFlag='" & s & "' And vd.Cust_DelFlag='" & s & "' ORDER BY st.Stor_Date " + sw
            stt = "Select ie.Money_ID as [رقم الأذن],ie.Money_Type_Arabic as [نوع الاذن], ie.Money_Date as [التاريخ], ie.Money_Price as [المبلغ], ie.Money_Reason as [البيــــان], ie.Money_Note as [ملاحظات] FROM Imp_Exp_Money as ie where ie.Cust_ID=" & cust_id & " And ie.Stor_ID=" & ze & " And ie.Money_DelFlag='" & s & "' ORDER BY ie.Money_Date " + sw

            'stt = "SELECT fo.Fetora_id as [رقم الفاتورة], cu.Cust_Name as [أسم العميل], fo.Fetora_date as [تاريخ الفاتورة ], fo.Fetora_blancebeforetype as [نوع قبل], fo.Fetora_blancebefore as [رصيد قبل], fo.Fetora_totalprice as [إجمالى الفاتورة], fo.Fetora_totalpay as [إجمالى المدفوع], fo.Fetora_blanceaftertype as [نوع بعد], fo.Fetora_blanceafter as [رصيد بعد] FROM Fetora as fo,Customers as cu where cu.Cust_DelFlag='" & s & "' And cu.Cust_ID=" & cust_id & " And fo.Fetora_custid=" & cust_id & " And cu.Cust_ID=fo.Fetora_custid ORDER BY fo.Fetora_id " + sw



            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)

            DataGridView3.DataSource = dt.DefaultView
            gcon.Close()
            dgvmoney()
            DataGridView3.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أدخل أسم العميل", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If

        Dim ds As String
        ds = "DESC"
        rrr = 1
        get_fetora(ds)
        get_order(ds)
        get_money(ds)
        rrrss = "notsearch"
        ComboBox1.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        TextBox2.Enabled = True
        Button1.Enabled = True

        Dim s As Boolean
        s = False
        Dim u_name As String
        u_name = ComboBox2.SelectedItem.ToString()
        If rs_Vendors.State = 1 Then rs_Vendors.Close()
        rs_Vendors.Open("Select * From Customers Where Cust_ID='" & cust_id & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Label19.Text = rs_Vendors("Cust_Blance_Type").Value
        Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
        rs_Vendors.Close()
        Label20.Visible = True
        Label16.Visible = True
        Label19.Visible = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If rrr = 1 Then


            Dim ds As String
            If ComboBox1.Text = "الأقدم تحت" Then

                ds = "DESC"
                get_fetora(ds)

            End If
            If ComboBox1.Text = "الأقدم فوق" Then

                ds = "ASC"
                get_fetora(ds)

            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("من فضلك أدخل رقم الفاتورة", MsgBoxStyle.Information, "تنبيه")
            TextBox2.Select()
            Exit Sub
        End If

        Dim ds As String
        ds = "DESC"
        rrr = 1

        fd = TextBox2.Text
        get_order_fetoraid(ds, fd)
        rrrss = "search"
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If rrrss = "notsearch" Then


            Dim ds As String
            If ComboBox3.Text = "الأقدم تحت" Then

                ds = "DESC"
                get_order(ds)

            End If
            If ComboBox3.Text = "الأقدم فوق" Then

                ds = "ASC"
                get_order(ds)

            End If
        End If
        If rrrss = "search" Then


            Dim ds As String
            If ComboBox3.Text = "الأقدم تحت" Then

                ds = "DESC"
                get_order_fetoraid(ds, fd)

            End If
            If ComboBox3.Text = "الأقدم فوق" Then

                ds = "ASC"
                get_order_fetoraid(ds, fd)

            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If rrr = 1 Then


            Dim ds As String
            If ComboBox4.Text = "الأقدم تحت" Then

                ds = "DESC"
                get_money(ds)

            End If
            If ComboBox4.Text = "الأقدم فوق" Then

                ds = "ASC"
                get_money(ds)

            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Cursor = Cursors.WaitCursor
        Dim rpt As New Rpt_Showcustomers_reviewstor
        Dim frm_rpt As New frm_rpt_showcustomers_reviewstor
        rdat.Clear()
        rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
        rcon.Open()
        Dim ssd As Boolean
        Dim ghf As Integer
        Dim sy3 As String
        Dim C_Name As String
        ssd = False

        rcomand = New SqlCommand("SELECT st.Stor_ID as storid,st.Fetora_id as storfatoriid, st.Stor_Type_Arabic as stortype, itm.Item_Name as storitemname, st.Stor_Date as stordate, st.Stor_Quntity_No as storquantity, st.Stor_Price_tin as storprice, st.Stor_Total_Price as stortotalprice, st.Stor_Payment as storpay FROM Stor as st, Customers as vd, Itemes as itm where st.Cust_ID=vd.Cust_ID And st.Item_ID=itm.Item_ID And st.Cust_ID=" & cust_id & " And vd.Cust_ID=" & cust_id & " And st.Stor_DelFlag='" & ssd & "' And vd.Cust_DelFlag='" & ssd & "' ORDER BY st.Stor_Date " + sw1, rcon)
        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        ghf = rs_StoreVend("Cust_Blance_Medin").Value
        sy3 = rs_StoreVend("Cust_Blance_Type").Value
        C_Name = rs_StoreVend("Cust_Name").Value
        rs_StoreVend.Close()
        rada = New SqlDataAdapter(rcomand)

        rada.Fill(rdat, "stor")
        rpt.SetDataSource(rdat)
        rpt.SetParameterValue("T_Name2", sy3)
        rpt.SetParameterValue("B_name", ghf)
        rpt.SetParameterValue("C_name", C_Name)
        frm_rpt.CrystalReportViewer1.ReportSource = rpt
        frm_rpt.CrystalReportViewer1.Refresh()
        frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
        Dim frm As New Form
        With frm
            .Controls.Add(frm_rpt.CrystalReportViewer1)
            .Text = "طباعة كل حركات المخزن لعميل"
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Cursor = Cursors.WaitCursor
        Dim rpt As New Rpt_Showcustomers_reviewmoney
        Dim frm_rpt As New frm_rpt_showcustomers_reviewmoney
        rdat.Clear()
        rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
        rcon.Open()
        Dim ssd As Boolean
        Dim ghf As Integer
        Dim sy3 As String
        Dim C_Name As String
        ssd = False
        Dim ze As Integer
        ze = 0
        rcomand = New SqlCommand("Select ie.Money_ID as moneyid,ie.Money_Type_Arabic as moneytype, ie.Money_Date as moneydate, ie.Money_Price as moneyprice FROM Imp_Exp_Money as ie where ie.Cust_ID=" & cust_id & " And ie.Stor_ID=" & ze & " And ie.Money_DelFlag='" & ssd & "' ORDER BY ie.Money_Date " + sw1, rcon)
        If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
        rs_StoreVend.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        ghf = rs_StoreVend("Cust_Blance_Medin").Value
        sy3 = rs_StoreVend("Cust_Blance_Type").Value
        C_Name = rs_StoreVend("Cust_Name").Value
        rs_StoreVend.Close()
        rada = New SqlDataAdapter(rcomand)

        rada.Fill(rdat, "money")
        rpt.SetDataSource(rdat)
        rpt.SetParameterValue("T_Name2", sy3)
        rpt.SetParameterValue("B_name", ghf)
        rpt.SetParameterValue("C_name", C_Name)
        frm_rpt.CrystalReportViewer1.ReportSource = rpt
        frm_rpt.CrystalReportViewer1.Refresh()
        frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
        Dim frm As New Form
        With frm
            .Controls.Add(frm_rpt.CrystalReportViewer1)
            .Text = "طباعة كل حركات الماليات لعميل"
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With

        Me.Cursor = Cursors.Default
    End Sub
End Class