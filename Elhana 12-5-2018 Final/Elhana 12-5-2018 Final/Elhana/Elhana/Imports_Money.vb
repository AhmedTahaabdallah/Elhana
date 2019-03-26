Imports ADODB
Public Class Imports_Money
    Public oe As Integer
    Public Order_Type As String
    Public month_id As Integer
    Public cust_idsearch As Integer
    Public cust_iddelete As Integer
    Public cust_id As Integer
    Public vend_idsearch As Integer
    Public vend_iddelete As Integer
    Public vend_id As Integer
    Public val_Old As Integer
    Public st_id As Integer
    Public ty As String
    Public val_Oldprice As Integer
    Public val_Oldpayment As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    Public cv As String
    Public isdafsave As String
    Public isdafedit As String
    Public isdafdelete As String


    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        RadioButton1.Checked = True
        TextBox1.Text = ""
        txt_mount.Text = ""
        txt_stat.Text = ""
        txt_Note.Text = ""
        ComboBox1.Text = ""
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        GroupBox3.Enabled = False
        btn_save.Enabled = True
        'ComboBox2.Enabled = True
        'ComboBox3.Enabled = False
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        DateTimePicker1.Enabled = True
        ComboBox1.Enabled = True
        DateTimePicker1.Select()
        Dim s As Boolean
        s = False
        Try
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(Money_ID) as df From Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label8.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("Money_ID").Value
                Label8.Text = u_name + 1
            End If
            
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Try
            Me.Dispose()
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
        If txt_mount.Text = "" Then
            MsgBox("من فضلك أدخل المبلغ", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_stat.Text = "" Then
            MsgBox("من فضلك أدخل بيان الأذن", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        If RadioButton1.Checked = True Then
            If ComboBox2.Text = "" Then
                MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
                ComboBox2.Select()
                Exit Sub
            End If
        End If
        If RadioButton3.Checked = True Then
            If ComboBox3.Text = "" Then
                MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
                ComboBox3.Select()
                Exit Sub
            End If
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
            Dim uk As Integer
            uk = 2
            Dim s As Boolean
            s = False
            Dim ss1 As Boolean
            ss1 = False
            'Dim val_storintotalprice As Integer
            Dim mo1 As Integer
            mo1 = Val(txt_mount.Text)
            Dim z As Integer
            z = 0
            Dim tt As Integer
            Dim ve As Integer
            Dim df As String
            Dim cu As Integer
            tt = cust_id
            'If ty = "cust" Then
            '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
            '    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '    If rs_Vendors.EOF Or rs_Vendors.BOF Then

            '    Else
            '        val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
            '    End If


            '    Dim ll As Integer
            '    ll = val_storintotalprice - mo1
            '    If ll < z Then
            '        MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
            '        txt_mount.Select()
            '        Exit Sub
            '    End If
            'End If

            'If rs_Customers.State = 1 Then rs_Customers.Close()
            'rs_Customers.Open("Select * From Imp_Exp_Money Where Money_DelFlag=" & s & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            'If rs_Customers.EOF Or rs_Customers.BOF Then

            'Dim u_name As Integer
            'Dim u_name1 As Integer
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAX(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'u_name = rs_StoreVend("Money_ID").Value
            'u_name1 = u_name + 1
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim rt As Integer
            rt = 0

            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_Customers.AddNew()
            'rs_Customers("Money_ID").Value = u_name1
            rs_Customers("Money_Type").Value = Order_Type
            rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
            rs_Customers("Money_Datestri").Value = dstri
            rs_Customers("Money_Price").Value = txt_mount.Text
            rs_Customers("Money_Reason").Value = txt_stat.Text
            rs_Customers("Money_Note").Value = txt_Note.Text
            rs_Customers("User_ID").Value = u_Id
            rs_Customers("User_ID_edit").Value = uk
            rs_Customers("User_ID_delete").Value = uk
            rs_Customers("Money_DelFlag").Value = s
            rs_Customers("Month_ID").Value = month_id

            If RadioButton3.Checked = True Then
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & vend_id & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    ve = rs_Vendors("Vend_ID").Value
                    cu = oe
                    df = rs_Vendors("Vend_Blance_d").Value
                    Dim tttt1 As Integer
                    tttt1 = rs_Vendors("Vend_Blance_Dein").Value
                    Dim jj As Integer

                    If df = "defulit" Then
                        jj = rs_Vendors("Vend_Blance_Dein").Value + Val(txt_mount.Text)
                        rs_Vendors("Vend_Blance_Dein").Value = jj
                        rs_Vendors.Update()
                    End If
                    If df = "notdefulit" Then

                        If tttt1 > mo1 Then
                            jj = tttt1 - Val(txt_mount.Text)
                            rs_Vendors("Vend_Blance_Dein").Value = jj
                            rs_Vendors.Update()
                        End If
                        If tttt1 <= mo1 Then
                            jj = Val(txt_mount.Text) - tttt1
                            rs_Vendors("Vend_Blance_Dein").Value = jj
                            rs_Vendors.Update()
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
            If RadioButton1.Checked = True Then
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    cu = rs_Vendors("Cust_ID").Value
                    ve = oe
                    df = rs_Vendors("Cust_Blance_d").Value
                    Dim tttt As Integer
                    tttt = rs_Vendors("Cust_Blance_Medin").Value
                    Dim jj As Integer

                    If df = "defulit" Then

                        If tttt >= mo1 Then
                            jj = tttt - Val(txt_mount.Text)
                            rs_Vendors("Cust_Blance_Medin").Value = jj
                            rs_Vendors.Update()
                        End If
                        If tttt < mo1 Then
                            jj = Val(txt_mount.Text) - tttt
                            rs_Vendors("Cust_Blance_Medin").Value = jj
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "ليه رصيد"
                            rs_Vendors("Cust_Blance_Medin").Value = jj
                            rs_Vendors("Cust_Blance_Type").Value = qw
                            rs_Vendors("Cust_Blance_d").Value = re
                            rs_Vendors.Update()

                        End If

                    End If
                    If df = "notdefulit" Then
                        jj = rs_Vendors("Cust_Blance_Medin").Value + Val(txt_mount.Text)
                        rs_Vendors("Cust_Blance_Medin").Value = jj
                        rs_Vendors.Update()
                    End If

                End If
            End If
            If RadioButton2.Checked = True Then
                cu = oe
                ve = oe
            End If
            rs_Customers("Money_Type_Arabic").Value = "وارد إلى الخزنة"
            rs_Customers("Cust_ID").Value = cu
            rs_Customers("Vend_ID").Value = ve
            rs_Customers("Stor_ID").Value = rt
            rs_Customers.Update()
            MsgBox("تم حفظ الأذن بنجاح", MsgBoxStyle.Information, "حفظ بيانات الأذن")
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            ComboBox1.Text = ""
            TextBox1.Text = ""
            btn_All.Enabled = True
            btn_etite.Enabled = True
            btn_delete.Enabled = True
            btn_add.Enabled = True
            GroupBox3.Enabled = True
            btn_save.Enabled = False

            Dim s8 As Boolean
            s8 = False
            Dim u_name8 As Integer

            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select* From Imp_Exp_Money Where Money_Type ='" & Order_Type & "' And Money_DelFlag='" & s8 & "' And Money_ID = (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            u_name8 = rs_StoreVend("Money_ID").Value

            Dim s3 As Boolean
            s3 = False
            Dim ne As Integer
            Dim ne22 As Integer
            Dim fin As Integer
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ne = rs_Store("Blance_Total").Value
            ne22 = rs_StoreVend("Money_Price").Value
            fin = ne + ne22

            'Dim u_name9 As Integer
            'Dim u_name21 As Integer
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreVend.Open("Select * From Money_Blance Where ID= (SELECT MAX(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'u_name9 = rs_StoreVend("ID").Value
            'u_name21 = u_name9 + 1
            s3 = False
            Dim mm As String
            mm = "حفظ وارد ماليات"
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_Store.AddNew()
            'rs_Store("ID").Value = u_name21
            rs_Store("Money_ID").Value = u_name8
            rs_Store("Blance_Total").Value = fin
            rs_Store("DelFlag").Value = s3
            rs_Store("Blance_State").Value = mm
            rs_Store("User_ID").Value = u_Id
            rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
            rs_Store.Update()

            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                'Exit Sub
            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Label27.Text = rs_Store("Blance_Total").Value
            End If
            rs_Store.Close()
            rs_StoreVend.Close()
            rs_Customers.Close()
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            ComboBox2.Visible = False

            RadioButton1.Checked = True
            TextBox1.Text = ""
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            ComboBox1.Text = ""
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            'ComboBox2.Enabled = True
            'ComboBox3.Enabled = False
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True
            DateTimePicker1.Enabled = True
            ComboBox1.Enabled = True
            DateTimePicker1.Select()
            'Dim s As Boolean
            s = False
            'Try
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(Money_ID) as df From Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label8.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("Money_ID").Value
                Label8.Text = u_name + 1
            End If

            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub Imports_Money_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_moneyimp_add").Value
            auth_show = rs_auth("User_moneyimp_show").Value
            auth_search = rs_auth("User_moneyimp_search").Value
            auth_edit = rs_auth("User_moneyimp_edit").Value
            auth_delete = rs_auth("User_moneyimp_delete").Value
            'auth_report = rs_auth("User_moneyimp_repot").Value
            If auth_add = ut Then
                'btn_add.Visible = False
                btn_save.Visible = False
            Else
                'btn_add.Visible = True
                btn_save.Visible = True
            End If
            If auth_show = ut Then
                btn_All.Visible = False
            Else
                btn_All.Visible = True
            End If
            If auth_search = ut Then
                GroupBox3.Visible = False
            Else
                GroupBox3.Visible = True
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
            auth_blance = rs_auth("User_blance").Value
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
                btn_All.Visible = True
                GroupBox3.Visible = True
                btn_etite.Visible = True
                btn_delete.Visible = True
                Label29.Visible = True
                Label27.Visible = True
                Label28.Visible = True
            End If
            oe = 1
            Order_Type = "Imp"
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            Label8.Text = "0"
            ComboBox1.Text = ""
            TextBox1.Text = ""
            DateTimePicker1.Enabled = True
            ComboBox1.Enabled = True
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
            Label19.Visible = False
            Label9.Visible = False


            RadioButton1.Checked = True
            TextBox1.Text = ""
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            ComboBox1.Text = ""
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            'ComboBox2.Enabled = True
            'ComboBox3.Enabled = False
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True
            DateTimePicker1.Enabled = True
            ComboBox1.Enabled = True
            DateTimePicker1.Select()
            Dim s As Boolean
            s = False
            'Try
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(Money_ID) as df From Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label8.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("Money_ID").Value
                Label8.Text = u_name + 1
            End If

            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = TextBox1.Text
            Dim order_Type As String
            order_Type = "Imp"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_DelFlag='" & s & "' And Money_ID=" & u_name & " And Money_Type='" & order_Type & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                TextBox1.Select()
                Exit Sub
            End If

            txt_Note.Text = rs_StoreVend("Money_Note").Value
            txt_stat.Text = rs_StoreVend("Money_Reason").Value
            txt_mount.Text = rs_StoreVend("Money_Price").Value
            val_Oldpayment = rs_StoreVend("Money_Price").Value
            val_Oldprice = rs_StoreVend("Money_Price").Value
            DateTimePicker1.Value = rs_StoreVend("Money_Date").Value
            Label8.Text = rs_StoreVend("Money_ID").Value
            val_Old = txt_mount.Text
            st_id = rs_StoreVend("Stor_ID").Value
            Dim z As Integer
            z = 0
            If st_id = z Then
                DateTimePicker1.Enabled = True
                ComboBox1.Enabled = True
            End If
            If st_id > z Then
                DateTimePicker1.Enabled = False
                ComboBox1.Enabled = False
            End If

            Dim s15 As Boolean
            s15 = False
            Dim gh5 As Integer
            gh5 = rs_StoreVend("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.Open("Select * From year_monthes Where Month_ID=" & gh5 & " And Month_DeleFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox1.Text = rs_U("Month_Name").Value
            month_id = gh5
            Dim uioc As Integer
            uioc = rs_StoreVend("Cust_ID").Value
            Dim uiov As Integer
            uiov = rs_StoreVend("Vend_ID").Value
            If uioc > oe Then
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & uioc & " And Cust_DelFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox2.Text = rs_Vendors("Cust_Name").Value
                RadioButton1.Checked = True
                cv = "cust"
                cust_id = rs_StoreVend("Cust_ID").Value
                cust_idsearch = rs_StoreVend("Cust_ID").Value
                cust_iddelete = rs_StoreVend("Cust_ID").Value
            End If
            If uiov > oe Then
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & uiov & " And Vend_DelFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox3.Text = rs_Vendors("Vend_Name").Value
                RadioButton3.Checked = True
                cv = "vend"
                vend_id = rs_StoreVend("Vend_ID").Value
                vend_idsearch = rs_StoreVend("Vend_ID").Value
                vend_iddelete = rs_StoreVend("Vend_ID").Value
            End If
           
            If uioc = oe And uiov = oe Then
                RadioButton2.Checked = True
                cv = "other"
            End If
            rs_StoreVend.Close()
            rs_U.Close()
            rs_Vendors.Close()
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            RadioButton3.Enabled = False
            ComboBox2.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        Dim z As Integer
        z = 0
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If txt_mount.Text = "" Or txt_mount.Text = z Then
            MsgBox("من فضلك أدخل المبلغ", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_stat.Text = "" Then
            MsgBox("من فضلك أدخل بيان الأذن", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
            Dim val_storin As Integer
            Dim x As Integer
            Dim val_addtostore As Integer
            Dim val_new As Integer
            Dim val_storintotalprice As Integer
            Dim val_Blance_Total2 As Integer
            Dim ss1 As Boolean
            ss1 = False
            Dim val_newpayment As Integer
            Dim xtotalprice As Integer
            'Dim val_addtototalprice As Integer
            Dim val_addtopayment As Integer
            Dim payme As Integer
            payme = Val(txt_mount.Text)
            val_newpayment = Val(txt_mount.Text)
            ' '' '' ''To sure about totalprice and payment not make blance of Customers < 0
            Dim tt As Integer

            If cv = "cust" Then
                tt = cust_idsearch
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If
                'val_newpayment = payme
                'If val_newpayment > val_Old Then
                '    val_addtototalprice = val_newpayment - val_Old
                '    Dim ll As Integer
                '    ll = val_storintotalprice - val_addtototalprice
                '    If ll < z Then
                '        MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '        txt_mount.Select()
                '        Exit Sub
                '    End If
                'End If
            End If

            If cv = "vend" Then
                tt = vend_idsearch
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                    isdafedit = rs_Vendors("Vend_Blance_d").Value
                End If
            End If
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                val_Blance_Total2 = rs_Store("Blance_Total").Value
            End If


            val_newpayment = payme
            If val_Old > val_newpayment Then
                val_addtopayment = val_Old - val_newpayment
                Dim ty77 As Integer
                ty77 = val_Blance_Total2 - val_addtopayment
                If ty77 < z Then
                    MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                    txt_mount.Select()
                    Exit Sub
                End If
            End If

            '' '' '' '' '' '' ''Update money blance

            val_new = Val(txt_mount.Text)
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            val_storin = rs_Store("Blance_Total").Value
            Dim vend_id1 As Integer
            val_newpayment = Val(txt_mount.Text)
            If val_Old < val_new Then
                val_addtostore = val_new - val_Old
                x = val_storin + val_addtostore
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & ss1 & "' And Stor_ID=" & st_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    vend_id1 = 0
                Else
                    vend_id1 = rs_StoreVend("Cust_ID").Value
                    rs_StoreVend("Stor_Payment").Value = val_new
                    rs_StoreVend.Update()
                End If

                'ss1 = False
                'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                'Else
                '    Dim jj As Integer
                '    jj = rs_Vendors("Cust_Blance_Medin").Value - val_addtostore
                '    rs_Vendors("Cust_Blance_Medin").Value = jj
                '    rs_Vendors.Update()
                'End If

                val_addtopayment = val_newpayment - val_Oldpayment
                If cv = "cust" Then
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
                'If st_id = z Then
                '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                '    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_ids & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                '    Else
                '        Dim jj As Integer
                '        jj = rs_Vendors("Cust_Blance_Medin").Value - val_addtostore
                '        rs_Vendors("Cust_Blance_Medin").Value = jj
                '        rs_Vendors.Update()
                '    End If
                'End If
            End If

            If val_Old = val_new Then
                x = val_storin
                'val_addtostore = 0
            End If

            If val_new < val_Old Then
                val_addtostore = val_Old - val_new
                If val_storin >= val_addtostore Then
                    x = val_storin - val_addtostore
                End If
                'If val_storin = val_addtostore Then
                '    x = 0
                'End If
                If val_storin < val_addtostore Then
                    MsgBox("المبلغ داخل الخزنة لا يكفى للتعديل", MsgBoxStyle.Information, "تنبيه")
                    txt_mount.Select()
                    Exit Sub
                End If

                ss1 = False
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & ss1 & "' And Stor_ID=" & st_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    vend_id1 = 0
                Else
                    vend_id1 = rs_StoreVend("Cust_ID").Value
                    rs_StoreVend("Stor_Payment").Value = val_new
                    rs_StoreVend.Update()
                End If

                'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                'Else
                '    Dim jj As Integer
                '    jj = rs_Vendors("Cust_Blance_Medin").Value + val_addtostore
                '    rs_Vendors("Cust_Blance_Medin").Value = jj
                '    rs_Vendors.Update()
                'End If

                'If st_id = z Then
                '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                '    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_ids & " And Cust_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                '    Else
                '        Dim jj As Integer
                '        jj = rs_Vendors("Cust_Blance_Medin").Value + val_addtostore
                '        rs_Vendors("Cust_Blance_Medin").Value = jj
                '        rs_Vendors.Update()
                '    End If
                'End If
                val_addtopayment = val_Oldpayment - val_newpayment
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
            End If

            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim f As Integer
            f = Label8.Text
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            rs_StoreVend("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
            rs_StoreVend("Money_Datestri").Value = dstri
            rs_StoreVend("Money_Price").Value = txt_mount.Text
            rs_StoreVend("Money_Reason").Value = txt_stat.Text
            rs_StoreVend("Money_Note").Value = txt_Note.Text
            rs_StoreVend("User_ID_edit").Value = u_Id
            rs_StoreVend("Month_ID").Value = month_id
            rs_StoreVend.Update()
            MsgBox("تم تعديل الأذن بنجاح", MsgBoxStyle.Information, "تعديل بيانات الأذن")
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            Label8.Text = "0"
            TextBox1.Text = ""
            ComboBox1.Text = ""

            Dim mm As String
            mm = "تعديل وارد"
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_Store.AddNew()
            rs_Store("Money_ID").Value = f
            rs_Store("Blance_Total").Value = x
            rs_Store("DelFlag").Value = ss1
            rs_Store("Blance_State").Value = mm
            rs_Store("User_ID").Value = u_Id
            rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
            rs_Store.Update()
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
            DateTimePicker1.Enabled = True
            ComboBox1.Enabled = True
            rs_Store.Close()
            rs_Vendors.Close()
            rs_StoreVend.Close()
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True


            RadioButton1.Checked = True
            TextBox1.Text = ""
            txt_mount.Text = ""
            txt_stat.Text = ""
            txt_Note.Text = ""
            ComboBox1.Text = ""
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            'ComboBox2.Enabled = True
            'ComboBox3.Enabled = False
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True
            DateTimePicker1.Enabled = True
            ComboBox1.Enabled = True
            DateTimePicker1.Select()
            Dim s As Boolean
            s = False
            'Try
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(Money_ID) as df From Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label8.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("Money_ID").Value
                Label8.Text = u_name + 1
            End If

            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

            
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If

        Try
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
            Dim z As Integer
            z = 0
            Dim g As String
            Dim mou1 As Integer
            Dim tt As Integer
            Dim val_storintotalprice As Integer
            'Dim val_newprice As Integer
            'Dim val_addtototalprice As Integer
            'Dim val_newpayment As Integer
            'Dim val_addtopayment As Integer
            mou1 = Val(txt_mount.Text)
            g = MsgBox("هل تريد حذف هذا الاذن ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

            If g = vbYes Then
                Dim s As Boolean
                s = False
                'Dim val_storintotalprice As Integer
                'Dim ctp As String
                Dim val_Blance_Total2 As Integer
                ' '' '' ''To sure about totalprice and payment not make blance of vendor < 0
                'Dim tt As Integer
                'tt = cust_iddelete

                'If RadioButton1.Checked = True Then



                '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                '    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '    If rs_Vendors.EOF Or rs_Vendors.BOF Then
                '        'ctp = "other"
                '    Else
                '        val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                '        'ctp = "cust"
                '    End If

                '    val_newprice = mou1
                '    If val_newprice > val_Oldprice Then
                '        val_addtototalprice = val_newprice - val_Oldprice
                '        Dim ll As Integer
                '        ll = val_storintotalprice - val_addtototalprice
                '        If ll < z Then
                '            MsgBox("رصيد العميل يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '            txt_mount.Select()
                '            Exit Sub
                '        End If
                '    End If

                'End If
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Store.EOF Or rs_Store.BOF Then

                Else
                    val_Blance_Total2 = rs_Store("Blance_Total").Value
                End If



                Dim ty3 As Integer
                ty3 = val_Blance_Total2 - mou1
                If ty3 < z Then
                    MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                    txt_mount.Select()
                    Exit Sub
                End If



                Dim r As Integer
                r = Label8.Text
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID=" & r & " And Money_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim oldprice As Integer
                oldprice = rs_StoreVend("Money_Price").Value

                If rs_Store.State = 1 Then rs_Store.Close()
                Dim s3 As Boolean
                s3 = False
                rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim val_in As Integer
                val_in = rs_Store("Blance_Total").Value


                If val_in >= oldprice Then

                    s = True
                    rs_StoreVend("Money_DelFlag").Value = s
                    rs_StoreVend("User_ID_delete").Value = u_Id
                    rs_StoreVend.Update()
                    MsgBox("تم حذف الحركة بنجاح", MsgBoxStyle.Information, "حذف بيانات توريده")
                    txt_mount.Text = ""
                    txt_stat.Text = ""
                    txt_Note.Text = ""
                    Label8.Text = "0"
                    TextBox1.Text = ""
                    ComboBox1.Text = ""

                    s = False
                    Dim fin As Integer
                    fin = val_in - oldprice
                    Dim mm As String
                    mm = "حذف وارد"
                    If rs_Store.State = 1 Then rs_Store.Close()
                    rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_Store.AddNew()
                    rs_Store("Money_ID").Value = r
                    rs_Store("Blance_Total").Value = fin
                    rs_Store("DelFlag").Value = s
                    rs_Store("Blance_State").Value = mm
                    rs_Store("User_ID").Value = u_Id
                    rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                    rs_Store.Update()

                    Dim vend_id As Integer
                    s3 = False
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & s3 & "' And Stor_ID=" & st_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                        vend_id = 0
                    Else
                        vend_id = rs_StoreVend("Cust_ID").Value
                        rs_StoreVend("Stor_Payment").Value = 0
                        rs_StoreVend.Update()
                    End If

                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    '    Dim jj As Integer
                    '    jj = rs_Vendors("Cust_Blance_Medin").Value + oldprice
                    '    rs_Vendors("Cust_Blance_Medin").Value = jj
                    '    rs_Vendors.Update()
                    'End If
                    If cv = "cust" Then
                        tt = cust_iddelete
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                            isdafdelete = rs_Vendors("Cust_Blance_d").Value
                        End If
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
                        tt = vend_iddelete
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
                If val_in < oldprice Then
                    MsgBox("رصيد الخزنة لا يسمح بارجاع المبلغ الوارد")

                End If



                'If st_id = z Then
                '    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                '    rs_Vendors.Open("Select * From Customers Where Cust_ID='" & cust_ids & "' And Cust_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                '    Else
                '        Dim jj As Integer
                '        jj = rs_Vendors("Cust_Blance_Medin").Value + oldprice
                '        rs_Vendors("Cust_Blance_Medin").Value = jj
                '        rs_Vendors.Update()
                '    End If
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
                DateTimePicker1.Enabled = True
                ComboBox1.Enabled = True
                rs_Vendors.Close()
                rs_Store.Close()
                rs_StoreVend.Close()
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                RadioButton3.Enabled = True


                RadioButton1.Checked = True
                TextBox1.Text = ""
                txt_mount.Text = ""
                txt_stat.Text = ""
                txt_Note.Text = ""
                ComboBox1.Text = ""
                'btn_All.Enabled = False
                'btn_etite.Enabled = False
                'btn_delete.Enabled = False
                'btn_add.Enabled = False
                'GroupBox3.Enabled = False
                btn_save.Enabled = True
                'ComboBox2.Enabled = True
                'ComboBox3.Enabled = False
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                RadioButton3.Enabled = True
                DateTimePicker1.Enabled = True
                ComboBox1.Enabled = True
                DateTimePicker1.Select()
                'Dim s As Boolean
                s = False
                'Try
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select Count(Money_ID) as df From Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                If rs_StoreVend("df").Value = 0 Then
                    Label8.Text = 1
                Else
                    Dim u_name As Integer
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("Select * From Imp_Exp_Money Where Money_ID= (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    u_name = rs_StoreVend("Money_ID").Value
                    Label8.Text = u_name + 1
                End If

                'Catch ex As Exception
                '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
                'End Try

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_mount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_mount.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub



    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub


    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        Frm_ShowImportsMoney.Show()
    End Sub

   

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Try
            ty = "cust"
            ComboBox2.Visible = True
            ComboBox3.Visible = False
            ComboBox2.Text = ""
            Label19.Visible = False
            Label9.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Try
            ty = "other"
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = False
            Label9.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try

            'If ty = "Vend" Then
            '    If rs_Cars.State = 1 Then rs_Cars.Close()
            '    Dim s As Boolean
            '    s = False
            '    rs_Cars.Open("Select * From Vendores Where Vend_DelFlag='" & s & "' ORDER BY Vend_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    ComboBox2.Items.Clear()
            '    Do While Not rs_Cars.EOF
            '        ComboBox2.Items.Add(rs_Cars("Vend_Name").Value)
            '        rs_Cars.MoveNext()
            '    Loop
            '    rs_Cars.Close()
            'End If

            'If ty = "Cust" Then
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From Customers Where Cust_DelFlag='" & s & "'  And Cust_visable='" & s & "' ORDER BY Cust_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox2.Items.Add(rs_Cars("Cust_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            'If ty = "Vend" Then
            '    Dim s As Boolean
            '    s = False
            '    Dim u_name As String
            '    u_name = ComboBox2.SelectedItem.ToString()
            '    If rs_Cars.State = 1 Then rs_Cars.Close()
            '    rs_Cars.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    vend_id = rs_Cars("Vend_ID").Value
            '    rs_Cars.Close()
            'End If

            'If ty = "Cust" Then
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox2.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            cust_id = rs_Cars("Cust_ID").Value
            Label19.Text = rs_Cars("Cust_Blance_Type").Value
            Label9.Text = rs_Cars("Cust_Blance_Medin").Value
            rs_Cars.Close()
            Label19.Visible = True
            Label9.Visible = True
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        Try

            'If ty = "Vend" Then
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.Open("Select * From Vendores Where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' ORDER BY Vend_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox3.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox3.Items.Add(rs_Cars("Vend_Name").Value)
                rs_Cars.MoveNext()
            Loop
            '    rs_Cars.Close()
            'End If

            'If ty = "Cust" Then
            'If rs_Cars.State = 1 Then rs_Cars.Close()
            'Dim s As Boolean
            's = False
            'rs_Cars.Open("Select * From Customers Where Cust_DelFlag='" & s & "' ORDER BY Cust_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            'ComboBox2.Items.Clear()
            'Do While Not rs_Cars.EOF
            '    ComboBox2.Items.Add(rs_Cars("Cust_Name").Value)
            '    rs_Cars.MoveNext()
            'Loop
            rs_Cars.Close()
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            'If ty = "Vend" Then
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox3.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            vend_id = rs_Cars("Vend_ID").Value
            Label19.Text = rs_Cars("Vend_Blance_Type").Value
            Label9.Text = rs_Cars("Vend_Blance_Dein").Value
            '    rs_Cars.Close()
            'End If

            'If ty = "Cust" Then
            'Dim s As Boolean
            's = False
            'Dim u_name As String
            'u_name = ComboBox2.SelectedItem.ToString()
            'If rs_Cars.State = 1 Then rs_Cars.Close()
            'rs_Cars.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            'cust_id = rs_Cars("Cust_ID").Value
            rs_Cars.Close()
            Label19.Visible = True
            Label9.Visible = True
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        Try
            ty = "vend"
            ComboBox2.Visible = False
            ComboBox3.Visible = True
            Label19.Visible = False
            Label9.Visible = False
            ComboBox3.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class