Imports ADODB
Public Class Frm_StoreVend
    Public vend_id As Integer
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
    Public val_Oldquntity As Integer
    Public month_id As Integer
    Public ite_id As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String
    Public cv As String
    Public stred As String
    Public strnoteed As String
    Public money_Typeed As String
    Public isdafsave As String
    Public isdafdelete As String
    Public isdafedit As String
    Public order_Typeedit As String
    Public oe As Integer


    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Vendores Where Vend_DelFlag='" & s & "'  And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "'  And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            vend_id = rs_Vendors("Vend_ID").Value

            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & vend_id & " And Vend_DelFlag='" & s & "'   And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Label16.Text = rs_Vendors("Vend_Blance_Dein").Value
            Label19.Text = rs_Vendors("Vend_Blance_Type").Value
            vend_state_defult = rs_Vendors("Vend_Blance_d").Value
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

    Private Sub ComboBox4_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.DropDown
        Try
            If rs_Citys.State = 1 Then rs_Citys.Close()
            Dim s As Boolean
            s = False
            rs_Citys.Open("Select * From Gover Where Gov_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox4.Items.Clear()
            Do While Not rs_Citys.EOF
                ComboBox4.Items.Add(rs_Citys("Gov_Name").Value)
                rs_Citys.MoveNext()
            Loop
            rs_Citys.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            Dim h As String
            h = ComboBox4.SelectedItem.ToString()
            If rs_Citys.State = 1 Then rs_Citys.Close()
            Dim s As Boolean
            s = False
            rs_Citys.Open("Select * From Gover Where Gov_Name='" & h & "' And Gov_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            gov_Name = rs_Citys("Gov_Name").Value
            rs_Citys.Close()
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
        Label31.Visible = False
        ComboBox8.Visible = False
        Label1.Visible = True
        ComboBox2.Visible = True
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox7.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        txt_quntity.Text = 0
        txt_quntity.Text = 0
        txt_Payment.Text = 0
        txt_price.Text = 0
        TextBox1.Text = ""
        Label15.Text = "-"
        Label16.Text = "-"
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        GroupBox3.Enabled = False
        btn_save.Enabled = True
        ComboBox2.Enabled = True
        Button2.Visible = True
        ComboBox7.Enabled = True
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
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If
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
        Dim jk As Integer

        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
            jk = 1
        End If

        Try
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
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
            Dim vendname As String
            vendname = ComboBox2.Text
            Dim val_storintotalprice As Integer
            Dim val_Blance_Total As Integer
            'Dim val_pay As Integer
            ' '' '' ''To sure about payment not make blance of Vendores < 0
            Dim z As Integer
            z = 0

            Dim tt As Integer
            tt = vend_id
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                isdafsave = rs_Vendors("Vend_Blance_d").Value
            End If
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                val_Blance_Total = rs_Store("Blance_Total").Value
            End If

            'val_pay = txt_Payment.Text
            'Dim ty2 As Integer
            'Dim ty5 As Integer
            'ty5 = val_storintotalprice + toprice
            'ty2 = ty5 - txt_Payment.Text

            'If ty2 < z Then
            '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
            '    txt_quntity.Select()
            '    Exit Sub
            'End If

            Dim mo As Integer
            mo = val_Blance_Total - payme
            If mo < z Then
                MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                txt_Payment.Select()
                Exit Sub
            End If
            '' '' '' ''--------------------------------------------
            Dim uk As Integer
            uk = 2
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd

            Dim sg As String
            sg = "رصيد مبدئى " + val_storintotalprice.ToString

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT COUNT(st.Stor_ID) as c_No FROM Stor as st, Vendores as vd where st.Vend_ID=" & vend_id & " And vd.Vend_ID=" & vend_id & " And st.Vend_ID=vd.Vend_ID And st.Stor_DelFlag='" & s & "' And vd.Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
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
            rs_StoreVend("Vend_ID").Value = vend_id
            rs_StoreVend("Cust_ID").Value = oe
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
            rs_StoreVend("Stor_Payment").Value = payme
            rs_StoreVend("Stor_Type_Arabic").Value = "وارد من"
            rs_StoreVend.Update()
            MsgBox("تم حفظ عملية الوارد بنجاح", MsgBoxStyle.Information, "حفظ بيانات الواردات")
            ComboBox2.Text = ""
            ComboBox1.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            ComboBox7.Text = ""
            txt_quntity.Text = 0
            txt_price.Text = 0
            txt_Payment.Text = 0
            txt_totalPrice.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"
            btn_All.Enabled = True
            btn_etite.Enabled = True
            btn_delete.Enabled = True
            btn_add.Enabled = True
            GroupBox3.Enabled = True
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

            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                isdafsave = rs_Vendors("Vend_Blance_d").Value
            End If

            ' ''ليه رصيد
            If isdafsave = "defulit" Then
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    Dim jj As Integer
                    jj = val_storintotalprice + toprice
                    rs_Vendors("Vend_Blance_Dein").Value = jj
                    rs_Vendors.Update()
                End If
            End If
            ' ''عليه رصيد
            If isdafsave = "notdefulit" Then
                If val_storintotalprice > toprice Then
                    Dim ff As Integer
                    ff = val_storintotalprice - toprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Vend_Blance_Dein").Value = ff
                        rs_Vendors.Update()
                    End If
                End If

                If val_storintotalprice <= toprice Then
                    Dim ff8 As Integer
                    ff8 = toprice - val_storintotalprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim re As String
                        Dim qw As String
                        re = "defulit"
                        qw = "ليه رصيد"
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Vend_Blance_Dein").Value = ff8
                        rs_Vendors("Vend_Blance_Type").Value = qw
                        rs_Vendors("Vend_Blance_d").Value = re
                        rs_Vendors.Update()
                    End If
                End If

            End If
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                isdafsave = rs_Vendors("Vend_Blance_d").Value
            End If
            If payme > 0 Then
                dy = DateTimePicker1.Value.Year
                dm = DateTimePicker1.Value.Month
                dd = DateTimePicker1.Value.Day
                dstri = dy + "-" + dm + "-" + dd
                order_Type = "Exp"
                Dim str As String
                str = "صادر إلى " + vendname
                Dim strnote As String
                Dim strn2 As String
                strn2 = u_name8
                strnote = "رقم أذن الوارد للمخزن " + strn2
                s = False
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Customers.AddNew()
                'rs_Customers("Money_ID").Value = u_name1
                rs_Customers("Money_Type").Value = order_Type
                rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
                rs_Customers("Money_Datestri").Value = dstri
                rs_Customers("Money_Price").Value = payme
                rs_Customers("Money_Reason").Value = str
                rs_Customers("Money_Note").Value = strnote
                rs_Customers("User_ID").Value = u_Id
                rs_Customers("Money_DelFlag").Value = s
                rs_Customers("Month_ID").Value = month_id
                rs_Customers("Stor_ID").Value = u_name8
                rs_Customers("Vend_ID").Value = vend_id
                rs_Customers("Cust_ID").Value = oe
                rs_Customers("Money_Type_Arabic").Value = "خارج من الخزنة"
                rs_Customers.Update()

                ' ''ليه رصيد
                If isdafsave = "defulit" Then
                    If val_storintotalprice >= payme Then
                        Dim ff As Integer
                        ff = val_storintotalprice - payme
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = ff
                            rs_Vendors.Update()
                        End If
                    End If

                    If val_storintotalprice < payme Then
                        Dim ff8 As Integer
                        ff8 = payme - val_storintotalprice
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "عليه رصيد"
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = ff8
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                End If
                ' ''عليه رصيد
                If isdafsave = "notdefulit" Then
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim jj As Integer
                        jj = val_storintotalprice + payme
                        rs_Vendors("Vend_Blance_Dein").Value = jj
                        rs_Vendors.Update()
                    End If


                End If

                s8 = False
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select* From Imp_Exp_Money Where Money_Type ='" & order_Type & "' And Money_DelFlag='" & s8 & "' And Money_ID = (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name8 = rs_StoreVend("Money_ID").Value

                'Dim s3 As Boolean
                's3 = False
                Dim ne9 As Integer
                Dim ne229 As Integer
                Dim fin9 As Integer
                s3 = False
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ne9 = rs_Store("Blance_Total").Value
                ne229 = rs_StoreVend("Money_Price").Value
                fin9 = ne9 - ne229

                Dim mm12 As String
                mm12 = "حفظ صادر"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = u_name8
                rs_Store("Blance_Total").Value = fin9
                rs_Store("DelFlag").Value = s3
                rs_Store("Blance_State").Value = mm12
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()

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
            rs_StoreVend.Close()
            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"
            Button2.Visible = False


            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox7.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            ComboBox2.Enabled = True
            Button2.Visible = True
            ComboBox7.Enabled = True
            ComboBox7.Select()
            'Dim s As Boolean
            s = False
            'Try


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
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

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
            order_Type = "Vend"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select * From Stor Where Stor_DelFlag='" & s & "' And Stor_ID=" & u_name & " And Stor_Type='" & order_Type & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                TextBox3.Select()
                Exit Sub
            End If

            'ComboBox5.Text = rs_StoreVend("Drivers_Name").Value
            'Dri_Name = rs_StoreVend("Drivers_Name").Value
            'ComboBox6.Text = rs_StoreVend("Stor_To").Value
            'ComboBox4.Text = rs_StoreVend("Stor_From").Value
            'gov_Name = rs_StoreVend("Stor_From").Value
            TextBox1.Text = rs_StoreVend("Stor_Note").Value
            txt_quntity.Text = rs_StoreVend("Stor_Quntity_No").Value
            val_Oldquntity = rs_StoreVend("Stor_Quntity_No").Value
            DateTimePicker1.Value = rs_StoreVend("Stor_Date").Value
            Label11.Text = rs_StoreVend("Stor_ID").Value
            txt_price.Text = rs_StoreVend("Stor_Price_tin").Value
            txt_totalPrice.Text = rs_StoreVend("Stor_Total_Price").Value
            val_Oldtotlprice = rs_StoreVend("Stor_Total_Price").Value
            txt_Payment.Text = rs_StoreVend("Stor_Payment").Value
            val_Oldpayment = rs_StoreVend("Stor_Payment").Value
            val_Old = Val(txt_quntity.Text)

            Dim uio As Integer
            uio = rs_StoreVend("Vend_ID").Value

            If uio > oe Then
                Dim s14 As Boolean
                s14 = False
                Dim gh4 As Integer
                gh4 = rs_StoreVend("Vend_ID").Value
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & gh4 & " And Vend_DelFlag='" & s14 & "'  And Vend_visable='" & s14 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox2.Text = rs_Vendors("Vend_Name").Value
                Label19.Text = rs_Vendors("Vend_Blance_Type").Value
                vend_state_defult = rs_Vendors("Vend_Blance_d").Value
                vend_idsearch = gh4
                vend_iddelete = gh4
                cv = "vend"
                Label16.Text = rs_Vendors("Vend_Blance_Dein").Value
                ComboBox8.Visible = False
                Label31.Visible = False
                ComboBox2.Enabled = False
                ComboBox2.Visible = True
                Label1.Visible = True
            End If
            If uio = oe Then
                Dim s1 As Boolean
                s1 = False
                Dim gh As Integer
                gh = rs_StoreVend("Cust_ID").Value
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & gh & " And Cust_DelFlag='" & s1 & "' And Cust_visable='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ComboBox8.Text = rs_Vendors("Cust_Name").Value
                Label19.Text = rs_Vendors("Cust_Blance_Type").Value
                cust_state_defult = rs_Vendors("Cust_Blance_d").Value
                cust_idsearch = gh
                cust_iddelete = gh
                cv = "cust"
                Label16.Text = rs_Vendors("Cust_Blance_Medin").Value
                ComboBox8.Visible = True
                Label31.Visible = True
                ComboBox8.Enabled = False
                ComboBox2.Visible = False
                Label1.Visible = False
            End If

            'Dim s1 As Boolean
            's1 = False
            'Dim gh As Integer
            'gh = rs_StoreVend("Vend_ID").Value
            'If rs_Vendors.State = 1 Then rs_Vendors.Close()
            'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & gh & " And Vend_DelFlag='" & s1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'ComboBox2.Text = rs_Vendors("Vend_Name").Value
            'vend_id = gh
            'vend_iddelete = gh

            'If rs_Vendors.State = 1 Then rs_Vendors.Close()
            'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & vend_id & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'Label16.Text = rs_Vendors("Vend_Blance_Dein").Value
            Dim s17 As Boolean
            s17 = False
            Dim gh1 As Integer
            Dim jh As String
            gh1 = rs_StoreVend("Item_ID").Value
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Itemes Where Item_ID=" & gh1 & " And Item_DelFlag='" & s17 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            jh = rs_Vendors("Item_Wight_Type").Value
            ComboBox7.Text = rs_Vendors("Item_Name").Value
            ite_id = gh1

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

            ComboBox7.Enabled = False
            rs_Vendors.Close()
            rs_StoreVend.Close()
            rs_U.Close()

            Button2.Visible = False
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
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
            Dim val_storin As Integer
            Dim val_Blance_Total As Integer
            Dim val_Blance_Total2 As Integer
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
            Dim rr As Integer
            Dim ss1 As Boolean
            ss1 = False
            Dim s As Boolean
            s = False


            order_Typeedit = "Vend"
            money_Typeed = "Exp"
            'If cv = "vend" Then


            'End If
            'If cv = "cust" Then
            '    order_Typeedit = "Cust"

            'End If
            Dim f As Integer
            f = Label11.Text
            Dim f1 As String
            f1 = Label11.Text
            Dim z As Integer
            z = 0
            Dim quentity As Integer
            quentity = Val(txt_quntity.Text)
            Dim toprice As Integer
            toprice = Val(txt_totalPrice.Text)
            Dim price As Integer
            price = Val(txt_price.Text)
            Dim payme As Integer
            payme = Val(txt_Payment.Text)

            ' '' '' ''To sure about totalprice and payment not make blance of vendor < 0
            Dim tt As Integer
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                val_Blance_Total2 = rs_Store("Blance_Total").Value
            End If
            If cv = "vend" Then
                tt = vend_idsearch
                'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                'Else
                '    val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                '    isdaf = rs_Vendors("Vend_Blance_d").Value
                'End If


                'val_newtotalprice = toprice
                'If val_newtotalprice < val_Oldtotlprice Then
                '    val_addtototalprice = val_Oldtotlprice - val_newtotalprice
                '    Dim ll As Integer
                '    ll = val_storintotalprice - val_addtototalprice
                '    If ll < z Then
                '        MsgBox("رصيد المورد يجب الايكون أقل من الصفر (قم بزيادة الكمية لزيادة أجمالى السعر )", MsgBoxStyle.Information, "تنبيه")
                '        txt_quntity.Select()
                '        Exit Sub
                '    End If
                'End If

                'Dim ade As Integer
                'Dim ades As Integer
                'val_newpayment = payme
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
                '    Dim ty8 As Integer
                '    ty8 = ades - val_addtopayment
                '    If ty8 < z Then
                '        MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '        txt_Payment.Select()
                '        Exit Sub
                '    End If
                'End If

                val_newpayment = payme
                If val_Oldpayment < val_newpayment Then
                    val_addtopayment = val_newpayment - val_Oldpayment
                    Dim ty98 As Integer
                    ty98 = val_Blance_Total2 - val_addtopayment
                    If ty98 < z Then
                        MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر  (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                        txt_Payment.Select()
                        Exit Sub
                    End If
                End If
            End If
            If cv = "cust" Then
                tt = cust_idsearch
                s = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If
                val_newpayment = payme
                If val_Oldpayment < val_newpayment Then
                    val_addtopayment = val_newpayment - val_Oldpayment
                    Dim ty98 As Integer
                    ty98 = val_Blance_Total2 - val_addtopayment
                    If ty98 < z Then
                        MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر  (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                        txt_Payment.Select()
                        Exit Sub
                    End If
                End If
            End If




            ' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''For stor
            val_new = quentity
            ss1 = False
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Store_Quentity Where Item_ID=" & ite_id & " And Store_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            val_storin = rs_Store("Store_Quntity").Value
            If val_Old < val_new Then
                val_addtostore = val_new - val_Old
                x = val_storin + val_addtostore
            End If
            If val_Old = val_new Then
                x = val_storin

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
                    MsgBox("الكمية داخل المخزن لا تكفى للتعديل", MsgBoxStyle.Information, "تنبيه")
                    txt_quntity.Select()
                    Exit Sub
                End If
            End If

            '' '' '' '' '' '' '' '' '' '' '' '' '' ''For  totalprice

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("Select * From Stor Where Stor_ID=" & f & " And Stor_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
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
                    If isdafedit = "defulit" Then
                        val_addtototalprice = val_newtotalprice - val_Oldtotlprice
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()
                    End If
                    If isdafedit = "notdefulit" Then
                        val_addtototalprice = val_newtotalprice - val_Oldtotlprice
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
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If

                    End If
                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    'rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                    'rs_Vendors.Update()
                    'End If

                End If


                If val_newtotalprice < val_Oldtotlprice Then
                    val_addtototalprice = val_Oldtotlprice - val_newtotalprice
                    If isdafedit = "defulit" Then
                        If val_addtototalprice <= val_storintotalprice Then
                            xtotalprice = val_storintotalprice - val_addtototalprice
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors.Update()
                        End If

                        If val_addtototalprice > val_storintotalprice Then
                            xtotalprice = val_addtototalprice - val_storintotalprice
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "عليه رصيد"
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                    If isdafedit = "notdefulit" Then
                        'val_addtototalprice = val_newtotalprice - val_Oldtotlprice
                        xtotalprice = val_storintotalprice + val_addtototalprice
                        rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                        rs_Vendors.Update()
                    End If

                    'Dim tow As Integer
                    'tow = val_storintotalprice - val_addtototalprice
                    'If tow < z Then
                    '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                    '    txt_quntity.Select()
                    '    Exit Sub
                    'End If



                    'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    'Else
                    'rs_Vendors("Vend_Blance_Dein").Value = xtotalprice
                    'rs_Vendors.Update()
                    'End If

                End If

                If val_Oldtotlprice = val_newtotalprice Then
                    rs_Vendors("Vend_Blance_Dein").Value = val_storintotalprice
                    rs_Vendors.Update()
                End If
            End If
            If cv = "cust" Then
                s = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If

                val_newtotalprice = toprice

                If val_Oldtotlprice < val_newtotalprice Then
                    val_addtototalprice = val_newtotalprice - val_Oldtotlprice
                    If isdafedit = "defulit" Then

                        If val_storintotalprice >= val_addtototalprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                xtotalprice = val_storintotalprice - val_addtototalprice
                                rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                                rs_Vendors.Update()
                            End If
                        End If
                        If val_storintotalprice < val_addtototalprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
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
                    End If
                    If isdafedit = "notdefulit" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            xtotalprice = val_storintotalprice + val_addtototalprice
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                    End If
                End If

                If val_newtotalprice < val_Oldtotlprice Then
                    val_addtototalprice = val_Oldtotlprice - val_newtotalprice
                    If isdafedit = "defulit" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            xtotalprice = val_storintotalprice + val_addtototalprice
                            rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                            rs_Vendors.Update()
                        End If
                    End If
                    If isdafedit = "notdefulit" Then
                        If val_storintotalprice > val_addtototalprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                xtotalprice = val_storintotalprice - val_addtototalprice
                                rs_Vendors("Cust_Blance_Medin").Value = xtotalprice
                                rs_Vendors.Update()
                            End If
                        End If
                        If val_storintotalprice <= val_addtototalprice Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
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
                    End If
                End If
            End If

            '' '' '' ''For payment and price in the export  and  blance tables
            ss1 = False
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & ss1 & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            val_Blance_Total = rs_Store("Blance_Total").Value

            If cv = "vend" Then
                'Dim jj3 As Integer
                'If rs_Vendors.State = 1 Then rs_Vendors.Close()
                'rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'If rs_Vendors.EOF Or rs_Vendors.BOF Then
                '    MsgBox("تأكد من المورد", MsgBoxStyle.Information, "تنبيه")
                '    ComboBox2.Select()
                '    Exit Sub
                'Else
                '    jj3 = rs_Vendors("Vend_Blance_Dein").Value
                'End If
                ss1 = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                    isdafedit = rs_Vendors("Vend_Blance_d").Value
                End If
            End If
            If cv = "cust" Then
                s = False
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    val_storintotalprice = rs_Vendors("Cust_Blance_Medin").Value
                    isdafedit = rs_Vendors("Cust_Blance_d").Value
                End If
            End If


            val_newpayment = payme
            If val_Oldpayment > val_newpayment Then
                val_addtopayment = val_Oldpayment - val_newpayment
                'Dim ty As Integer
                'ty = jj3 - val_addtopayment
                'If ty < 0 Then
                '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                '    txt_quntity.Select()
                '    Exit Sub
                'End If
                xpayment = val_Blance_Total + val_addtopayment

                If cv = "vend" Then
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


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    rr = 0
                Else
                    rr = rs_StoreVend("Money_ID").Value
                    rs_StoreVend("Money_Price").Value = val_newpayment
                    rs_StoreVend.Update()
                End If

                Dim mm12 As String
                mm12 = "تعديل صادر"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = rr
                rs_Store("Blance_Total").Value = xpayment
                rs_Store("DelFlag").Value = ss1
                rs_Store("Blance_State").Value = mm12
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()
            End If

            If val_newpayment > val_Oldpayment Then
                val_addtopayment = val_newpayment - val_Oldpayment

                If val_Blance_Total >= val_addtopayment Then
                    xpayment = val_Blance_Total - val_addtopayment
                End If
                'If val_Blance_Total = val_addtopayment Then
                '    xpayment = 0
                'End If

                If val_Blance_Total < val_addtopayment Then
                    MsgBox("المبلغ داخل الخزنة لا يكفى للتعديل", MsgBoxStyle.Information, "تنبيه")
                    txt_Payment.Select()
                    Exit Sub
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
                        rs_Customers("Money_Type").Value = money_Typeed
                        rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
                        rs_Customers("Money_Datestri").Value = dstri
                        rs_Customers("Money_Price").Value = val_newpayment

                        'Dim str As String

                        'Dim strnote As String

                        If cv = "vend" Then
                            stred = "صادر إلى " + ComboBox2.Text
                            strnoteed = "رقم أذن الوارد للمخزن " + f1
                            rs_Customers("Cust_ID").Value = oe
                            rs_Customers("Vend_ID").Value = tt
                        End If
                        If cv = "cust" Then
                            stred = "صادر إلى " + ComboBox8.Text
                            strnoteed = "رقم أذن الوارد للمخزن " + f1
                            rs_Customers("Cust_ID").Value = tt
                            rs_Customers("Vend_ID").Value = oe
                        End If
                        rs_Customers("Money_Type_Arabic").Value = "خارج من الخزنة"
                        rs_Customers("Money_Reason").Value = stred
                        rs_Customers("Money_Note").Value = strnoteed
                        rs_Customers("User_ID").Value = u_Id
                        rs_Customers("Money_DelFlag").Value = ss1
                        rs_Customers("Month_ID").Value = month_id
                        rs_Customers("Stor_ID").Value = f
                        'rs_Customers("Vend_ID").Value = vend_id

                        rs_Customers.Update()
                    Else
                        rs_StoreVend("Money_Price").Value = val_newpayment
                        rs_StoreVend.Update()
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
                If cv = "cust" Then
                    s = False
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
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


                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Imp_Exp_Money Where Stor_ID=" & f & " And Money_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                    rr = 0
                Else
                    rr = rs_StoreVend("Money_ID").Value
                    rs_StoreVend("Money_Price").Value = val_newpayment
                    rs_StoreVend.Update()
                End If

                Dim mm122 As String
                mm122 = "تعديل صادر"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = rr
                rs_Store("Blance_Total").Value = xpayment
                rs_Store("DelFlag").Value = ss1
                rs_Store("Blance_State").Value = mm122
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
            'rs_StoreVend("car_Id").Value = car_id
            'rs_StoreVend("Stor_From").Value = gov_Name
            'rs_StoreVend("Stor_To").Value = ComboBox6.Text
            'rs_StoreVend("Drivers_Name").Value = Dri_Name
            rs_StoreVend("User_ID_edit").Value = u_Id
            'rs_StoreVend("Vend_ID").Value = vend_id
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
            MsgBox("تم تعديل عملية الوارد بنجاح", MsgBoxStyle.Information, "تعديل بيانات الواردات")
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            ComboBox1.Text = ""
            ComboBox4.Text = ""
            ComboBox7.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = 0
            txt_price.Text = 0
            txt_Payment.Text = 0
            txt_totalPrice.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"

            ' '' '' '' '' '' ''Update store quntity
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & ss1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                Dim mm9 As String
                mm9 = "تعديل وارد"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Store_Quentity", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()

                rs_Store("Store_Quntity").Value = x
                rs_Store("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store("Stor_ID").Value = f
                rs_Store("Store_DelFlag").Value = ss1
                rs_Store("User_ID").Value = u_Id
                rs_Store("Stroe_State").Value = mm9
                rs_Store("Item_ID").Value = ite_id
                rs_Store.Update()
            Else
                Dim mm8 As String
                mm8 = "تعديل وارد"
                rs_StoreVend("Store_Quntity").Value = x
                rs_StoreVend("Store_Date").Value = Date.Now.Date.ToShortDateString()
                rs_StoreVend("Stor_ID").Value = f
                rs_StoreVend("Store_DelFlag").Value = ss1
                rs_StoreVend("User_ID").Value = u_Id
                rs_StoreVend("Stroe_State").Value = mm8
                rs_StoreVend("Item_ID").Value = ite_id
                rs_StoreVend.Update()
            End If

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
            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            rs_Store.Close()
            rs_Vendors.Close()
            rs_StoreVend.Close()
            'rs_Customers.Close()
            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"
            ComboBox2.Enabled = True
            ComboBox7.Enabled = True



            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox7.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            ComboBox2.Enabled = True
            Button2.Visible = True
            ComboBox7.Enabled = True
            ComboBox7.Select()
            'Dim s As Boolean
            s = False
            'Try


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
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

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

            If cv = "vend" Then
                tt = vend_iddelete
            End If
            If cv = "cust" Then
                tt = cust_iddelete
            End If
            'If cn1.State = 1 Then cn1.Close()
            'cn1.Open()
            Dim val_storintotalprice As Integer
            Dim z As Integer
            z = 0
            Dim g As String
            g = MsgBox("هل تريد حذف هذا الاذن ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

            If g = vbYes Then


                Dim g1 As String
                g1 = MsgBox("هل تريد حذف المبلغ المدفوع أيضا ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g1 = vbYes Then
                    st25 = "yesDeletepayment"
                    If cv = "vend" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                            isdafdelete = rs_Vendors("Vend_Blance_d").Value
                        End If
                        'Dim ll As Integer
                        'Dim ll77 As Integer
                        'll77 = val_storintotalprice + val_Oldpayment
                        'll = ll77 - val_Oldtotlprice
                        'If ll < z Then
                        '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                        '    txt_quntity.Select()
                        '    Exit Sub
                        'End If
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
                    End If
                    
                    
                Else
                    st25 = "noDeletepayment"
                    If cv = "vend" Then
                        s = False
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                            isdafdelete = rs_Vendors("Vend_Blance_d").Value
                        End If
                        'Dim ll33 As Integer
                        'll33 = val_storintotalprice - val_Oldtotlprice
                        'If ll33 < z Then
                        '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر", MsgBoxStyle.Information, "تنبيه")
                        '    txt_quntity.Select()
                        '    Exit Sub
                        'End If
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

                If val_in >= ne22 Then

                    s = True
                    rs_StoreVend("Stor_DelFlag").Value = s
                    rs_StoreVend("User_ID_delete").Value = u_Id
                    rs_StoreVend.Update()
                    MsgBox("تم حذف الحركة بنجاح", MsgBoxStyle.Information, "حذف بيانات توريده")
                    ComboBox2.Text = ""
                    ComboBox1.Text = ""
                    ComboBox3.Text = ""
                    ComboBox4.Text = ""
                    ComboBox5.Text = ""
                    ComboBox6.Text = ""
                    txt_quntity.Text = 0
                    txt_price.Text = 0
                    txt_Payment.Text = 0
                    txt_totalPrice.Text = 0
                    TextBox1.Text = ""
                    Label11.Text = "0"
                    ComboBox7.Text = ""

                    Label31.Visible = False
                    ComboBox8.Visible = False
                    Label1.Visible = True
                    ComboBox2.Visible = True
                    rs_Store.Close()
                    rs_StoreVend.Close()
                    rs_Vendors.Close()
                    Label15.Text = "-"
                    Label16.Text = "-"
                    Label11.Text = "-"
                    ComboBox2.Enabled = True
                    ComboBox7.Enabled = True
                    Dim fin As Integer
                    fin = val_in - ne22
                    s3 = False
                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.Open("SELECT * FROM Store_Quentity where Item_ID=" & ite_id & " And Store_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_StoreVend.EOF Or rs_StoreVend.BOF Then
                        Dim mm As String
                        mm = "حذف وارد"
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
                        mm = "حذف وارد"

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
                        fin2 = val_Blance_Total + val_Oldpayment
                        Dim mm1 As String
                        mm1 = "حذف صادر"
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

                        If cv = "vend" Then
                            

                            ' ''ليه رصيد
                            If isdafdelete = "defulit" Then
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim jj As Integer
                                    jj = val_storintotalprice + val_Oldpayment
                                    rs_Vendors("Vend_Blance_Dein").Value = jj
                                    rs_Vendors.Update()
                                End If


                            End If
                            ' ''عليه رصيد
                            If isdafdelete = "notdefulit" Then

                                If val_storintotalprice > val_Oldpayment Then
                                    Dim ff As Integer
                                    ff = val_storintotalprice - val_Oldpayment
                                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                    Else
                                        'Dim jj As Integer
                                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                        rs_Vendors("Vend_Blance_Dein").Value = ff
                                        rs_Vendors.Update()
                                    End If
                                End If

                                If val_storintotalprice <= val_Oldpayment Then
                                    Dim ff8 As Integer
                                    ff8 = val_Oldpayment - val_storintotalprice
                                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                    Else
                                        Dim re As String
                                        Dim qw As String
                                        re = "defulit"
                                        qw = "ليه رصيد"
                                        'Dim jj As Integer
                                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                        rs_Vendors("Vend_Blance_Dein").Value = ff8
                                        rs_Vendors("Vend_Blance_Type").Value = qw
                                        rs_Vendors("Vend_Blance_d").Value = re
                                        rs_Vendors.Update()
                                    End If
                                End If

                            End If

                        End If
                        If cv = "cust" Then
                            ' ''عليه رصيد
                            If isdafdelete = "defulit" Then

                                If val_storintotalprice >= val_Oldpayment Then
                                    s = False
                                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                    Else
                                        Dim ff8 As Integer
                                        ff8 = val_storintotalprice - val_Oldpayment
                                        rs_Vendors("Cust_Blance_Medin").Value = ff8
                                        rs_Vendors.Update()
                                    End If
                                End If
                                If val_storintotalprice < val_Oldpayment Then
                                    s = False
                                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                    rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                    Else
                                        Dim ff8 As Integer
                                        ff8 = val_Oldpayment - val_storintotalprice
                                        Dim re As String
                                        Dim qw As String
                                        re = "notdefulit"
                                        qw = "ليه رصيد"
                                        rs_Vendors("Cust_Blance_Type").Value = qw
                                        rs_Vendors("Cust_Blance_d").Value = re
                                        rs_Vendors("Cust_Blance_Medin").Value = ff8
                                        rs_Vendors.Update()
                                    End If
                                End If
                               
                            End If
                            ' ''ليه رصيد
                            If isdafdelete = "notdefulit" Then
                                s = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim ff8 As Integer
                                    ff8 = val_storintotalprice + val_Oldpayment
                                    rs_Vendors("Cust_Blance_Medin").Value = ff8
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
                            If val_storintotalprice >= val_Oldtotlprice Then
                                Dim ff As Integer
                                ff = val_storintotalprice - val_Oldtotlprice
                                s3 = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    'Dim jj As Integer
                                    'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                    rs_Vendors("Vend_Blance_Dein").Value = ff
                                    rs_Vendors.Update()
                                End If
                            End If

                            If val_storintotalprice < val_Oldtotlprice Then
                                Dim ff8 As Integer
                                ff8 = val_Oldtotlprice - val_storintotalprice
                                s3 = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim re As String
                                    Dim qw As String
                                    re = "notdefulit"
                                    qw = "عليه رصيد"
                                    'Dim jj As Integer
                                    'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                                    rs_Vendors("Vend_Blance_Dein").Value = ff8
                                    rs_Vendors("Vend_Blance_Type").Value = qw
                                    rs_Vendors("Vend_Blance_d").Value = re
                                    rs_Vendors.Update()
                                End If
                            End If
                        End If
                        ' ''عليه رصيد
                        If isdafdelete = "notdefulit" Then
                            s3 = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim jj As Integer
                                jj = val_storintotalprice + val_Oldtotlprice
                                rs_Vendors("Vend_Blance_Dein").Value = jj
                                rs_Vendors.Update()
                            End If

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

                        ' ''عليه رصيد
                        If isdafdelete = "defulit" Then
                            s = False
                            If rs_Vendors.State = 1 Then rs_Vendors.Close()
                            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rs_Vendors.EOF Or rs_Vendors.BOF Then

                            Else
                                Dim ff8 As Integer
                                ff8 = val_storintotalprice + val_Oldtotlprice
                                rs_Vendors("Cust_Blance_Medin").Value = ff8
                                rs_Vendors.Update()
                            End If

                            
                        End If
                        ' ''ليه رصيد
                        If isdafdelete = "notdefulit" Then

                            If val_storintotalprice > val_Oldtotlprice Then
                                s = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim ff8 As Integer
                                    ff8 = val_storintotalprice - val_Oldtotlprice
                                    rs_Vendors("Cust_Blance_Medin").Value = ff8
                                    rs_Vendors.Update()
                                End If
                            End If
                            If val_storintotalprice <= val_Oldtotlprice Then
                                s = False
                                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                                rs_Vendors.Open("Select * From Customers Where Cust_ID=" & tt & " And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                                Else
                                    Dim ff8 As Integer
                                    ff8 = val_Oldtotlprice - val_storintotalprice
                                    Dim re As String
                                    Dim qw As String
                                    re = "defulit"
                                    qw = "عليه رصيد"
                                    rs_Vendors("Cust_Blance_Type").Value = qw
                                    rs_Vendors("Cust_Blance_d").Value = re
                                    rs_Vendors("Cust_Blance_Medin").Value = ff8
                                    rs_Vendors.Update()
                                    rs_Vendors.Update()
                                End If
                            End If
                        End If

                End If
                
                End If
                If val_in < ne22 Then
                    MsgBox("رصيد المخزن لا يسمح بارجاع شحنة الوارد", MsgBoxStyle.Information, "تحذير")
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
                
            End If




            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox7.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            ComboBox2.Enabled = True
            Button2.Visible = True
            ComboBox7.Enabled = True
            ComboBox7.Select()
            'Dim s As Boolean
            s = False
            'Try


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
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

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

    Private Sub Frm_StoreVend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_storvend_add").Value
            auth_show = rs_auth("User_storvend_show").Value
            auth_search = rs_auth("User_storvend_search").Value
            auth_edit = rs_auth("User_storvend_edit").Value
            auth_delete = rs_auth("User_storvend_delete").Value
            'auth_report = rs_auth("User_storvend_report").Value
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
                btn_All.Visible = True
                GroupBox3.Visible = True
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
            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox7.Text = ""
            ComboBox2.Text = ""
            ComboBox1.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = ""
            TextBox1.Text = ""

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

            Label31.Visible = False
            ComboBox8.Visible = False
            Label1.Visible = True
            ComboBox2.Visible = True
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox7.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            txt_quntity.Text = 0
            txt_quntity.Text = 0
            txt_Payment.Text = 0
            txt_price.Text = 0
            TextBox1.Text = ""
            Label15.Text = "-"
            Label16.Text = "-"
            'btn_All.Enabled = False
            'btn_etite.Enabled = False
            'btn_delete.Enabled = False
            'btn_add.Enabled = False
            'GroupBox3.Enabled = False
            btn_save.Enabled = True
            ComboBox2.Enabled = True
            Button2.Visible = True
            ComboBox7.Enabled = True
            ComboBox7.Select()
            Dim s As Boolean
            s = False
            'Try


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
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            'End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        Frm_ShowStoreVend.Show()
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

            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Customers Where Cust_ID=" & cust_id & " And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Label16.Text = rs_Vendors("Cust_Blance_Medin").Value

            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If
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
            Dim vendname As String
            vendname = ComboBox2.Text
            Dim val_storintotalprice As Integer
            Dim val_Blance_Total As Integer
            'Dim val_pay As Integer
            Dim z As Integer
            z = 0

            ' '' '' ''To sure about payment not make blance of Vendores < 0
            Dim tt As Integer
            tt = vend_id
            vendid_cdid = vend_id
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                isdafsave = rs_Vendors("Vend_Blance_d").Value
            End If
            If rs_Store.State = 1 Then rs_Store.Close()
            rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & s & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Store.EOF Or rs_Store.BOF Then

            Else
                val_Blance_Total = rs_Store("Blance_Total").Value
            End If

            'val_pay = txt_Payment.Text
            'Dim ty2 As Integer
            'Dim ty5 As Integer
            'ty5 = val_storintotalprice + toprice
            'ty2 = ty5 - txt_Payment.Text

            'If ty2 < z Then
            '    MsgBox("رصيد المورد يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
            '    txt_quntity.Select()
            '    Exit Sub
            'End If

            Dim mo As Integer
            mo = val_Blance_Total - payme
            If mo < z Then
                MsgBox("رصيد الخزنة يجب الايكون أقل من الصفر (قم بأنقاص المبلغ المدفوع) ", MsgBoxStyle.Information, "تنبيه")
                txt_Payment.Select()
                Exit Sub
            End If
            '' '' '' ''--------------------------------------------
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Stor", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs_StoreVend.AddNew()
            'rs_StoreVend("Stor_ID").Value = u_name1
            'rs_StoreVend("car_Id").Value = car_id
            'rs_StoreVend("Stor_From").Value = gov_Name
            'rs_StoreVend("Stor_To").Value = ComboBox6.Text
            'rs_StoreVend("Drivers_Name").Value = Dri_Name
            rs_StoreVend("User_ID").Value = u_Id
            rs_StoreVend("Vend_ID").Value = vend_id
            rs_StoreVend("Cust_ID").Value = oe
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
            rs_StoreVend("Stor_Payment").Value = payme
            rs_StoreVend("Stor_Type_Arabic").Value = "وارد من"
            rs_StoreVend.Update()
            MsgBox("تم حفظ عملية الوارد بنجاح", MsgBoxStyle.Information, "حفظ بيانات الواردات")
            ComboBox2.Text = ""
            ComboBox1.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            ComboBox7.Text = ""
            txt_quntity.Text = 0
            txt_price.Text = 0
            txt_Payment.Text = 0
            txt_totalPrice.Text = 0
            TextBox1.Text = ""
            Label11.Text = "0"
            btn_All.Enabled = True
            btn_etite.Enabled = True
            btn_delete.Enabled = True
            btn_add.Enabled = True
            GroupBox3.Enabled = True
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
            sid_Vend_Cdid = rs_StoreVend("Stor_ID").Value
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

           
            ' ''ليه رصيد
            If isdafsave = "defulit" Then
                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Vendors.EOF Or rs_Vendors.BOF Then

                Else
                    Dim jj As Integer
                    jj = val_storintotalprice + toprice
                    rs_Vendors("Vend_Blance_Dein").Value = jj
                    rs_Vendors.Update()
                End If
            End If
            ' ''عليه رصيد
            If isdafsave = "notdefulit" Then
                If val_storintotalprice > toprice Then
                    Dim ff As Integer
                    ff = val_storintotalprice - toprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Vend_Blance_Dein").Value = ff
                        rs_Vendors.Update()
                    End If
                End If

                If val_storintotalprice < toprice Then
                    Dim ff8 As Integer
                    ff8 = toprice - val_storintotalprice
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim re As String
                        Dim qw As String
                        re = "defulit"
                        qw = "ليه رصيد"
                        'Dim jj As Integer
                        'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                        rs_Vendors("Vend_Blance_Dein").Value = ff8
                        rs_Vendors("Vend_Blance_Type").Value = qw
                        rs_Vendors("Vend_Blance_d").Value = re
                        rs_Vendors.Update()
                    End If
                End If

            End If
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Vendors.EOF Or rs_Vendors.BOF Then

            Else
                val_storintotalprice = rs_Vendors("Vend_Blance_Dein").Value
                isdafsave = rs_Vendors("Vend_Blance_d").Value
            End If


            If payme > 0 Then
                dy = DateTimePicker1.Value.Year
                dm = DateTimePicker1.Value.Month
                dd = DateTimePicker1.Value.Day
                dstri = dy + "-" + dm + "-" + dd
                order_Type = "Exp"
                Dim str As String
                str = "صادر إلى " + vendname
                Dim strnote As String
                Dim strn2 As String
                strn2 = u_name8
                strnote = "رقم أذن الوارد للمخزن " + strn2
                s = False
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Imp_Exp_Money", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Customers.AddNew()
                'rs_Customers("Money_ID").Value = u_name1
                rs_Customers("Money_Type").Value = order_Type
                rs_Customers("Money_Date").Value = DateTimePicker1.Value.ToShortDateString()
                rs_Customers("Money_Datestri").Value = dstri
                rs_Customers("Money_Price").Value = payme
                rs_Customers("Money_Reason").Value = str
                rs_Customers("Money_Note").Value = strnote
                rs_Customers("User_ID").Value = u_Id
                rs_Customers("Money_DelFlag").Value = s
                rs_Customers("Month_ID").Value = month_id
                rs_Customers("Stor_ID").Value = u_name8
                rs_Customers("Vend_ID").Value = vend_id
                rs_Customers("Cust_ID").Value = oe
                rs_Customers("Money_Type_Arabic").Value = "خارج من الخزنة"
                rs_Customers.Update()
               
                ' ''ليه رصيد
                If isdafsave = "defulit" Then
                    If val_storintotalprice > payme Then
                        Dim ff As Integer
                        ff = val_storintotalprice - payme
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = ff
                            rs_Vendors.Update()
                        End If
                    End If

                    If val_storintotalprice < payme Then
                        Dim ff8 As Integer
                        ff8 = payme - val_storintotalprice
                        If rs_Vendors.State = 1 Then rs_Vendors.Close()
                        rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If rs_Vendors.EOF Or rs_Vendors.BOF Then

                        Else
                            Dim re As String
                            Dim qw As String
                            re = "notdefulit"
                            qw = "عليه رصيد"
                            'Dim jj As Integer
                            'jj = rs_Vendors("Cust_Blance_Medin").Value + toprice
                            rs_Vendors("Vend_Blance_Dein").Value = ff8
                            rs_Vendors("Vend_Blance_Type").Value = qw
                            rs_Vendors("Vend_Blance_d").Value = re
                            rs_Vendors.Update()
                        End If
                    End If
                End If
                ' ''عليه رصيد
                If isdafsave = "notdefulit" Then
                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID=" & tt & " And Vend_DelFlag='" & s3 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    If rs_Vendors.EOF Or rs_Vendors.BOF Then

                    Else
                        Dim jj As Integer
                        jj = val_storintotalprice + payme
                        rs_Vendors("Vend_Blance_Dein").Value = jj
                        rs_Vendors.Update()
                    End If


                End If

                s8 = False
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select* From Imp_Exp_Money Where Money_Type ='" & order_Type & "' And Money_DelFlag='" & s8 & "' And Money_ID = (SELECT MAx(Money_ID)  FROM Imp_Exp_Money)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name8 = rs_StoreVend("Money_ID").Value

                'Dim s3 As Boolean
                's3 = False
                Dim ne9 As Integer
                Dim ne229 As Integer
                Dim fin9 As Integer
                s3 = False
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Select * From Money_Blance Where DelFlag='" & s3 & "' And ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                ne9 = rs_Store("Blance_Total").Value
                ne229 = rs_StoreVend("Money_Price").Value
                fin9 = ne9 - ne229

                Dim mm12 As String
                mm12 = "حفظ صادر"
                If rs_Store.State = 1 Then rs_Store.Close()
                rs_Store.Open("Money_Blance", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.AddNew()
                'rs_Store("ID").Value = u_name21
                rs_Store("Money_ID").Value = u_name8
                rs_Store("Blance_Total").Value = fin9
                rs_Store("DelFlag").Value = s3
                rs_Store("Blance_State").Value = mm12
                rs_Store("User_ID").Value = u_Id
                rs_Store("Blance_Date").Value = Date.Now.Date.ToShortDateString()
                rs_Store.Update()

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
            rs_StoreVend.Close()
            Label15.Text = "-"
            Label16.Text = "-"
            Label11.Text = "-"
            Button2.Visible = False
            frm_StoreVend_Cdid.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class