Imports ADODB
Public Class Frm_Customers

    Public C_name As String
    Public C_id As String


    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        txt_Name.Text = ""
        txt_Phone.Text = "014"
        txt_Address.Text = "لا يوجد عنوان"
        txt_Note.Text = ""
        TextBox2.Text = ""
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        ComboBox2.Visible = True
        ComboBox2.Text = "عليه رصيد"
        Label7.Visible = False
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        ComboBox1.Enabled = False
        btn_save.Enabled = True
        TextBox2.Enabled = True
        txt_Name.Select()
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_Customers.State = 1 Then rs_Customers.Close()
            Dim s As Boolean
            s = False
            rs_Customers.Open("Select * From Customers Where Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_Customers.EOF
                ComboBox1.Items.Add(rs_Customers("Cust_Name").Value)
                rs_Customers.MoveNext()
            Loop
            rs_Customers.Close()
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
            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Select * From Customers Where Cust_Name='" & u_name & "' And Cust_DelFlag='" & s & "' And Cust_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            txt_Name.Text = rs_Customers("Cust_Name").Value
            txt_Phone.Text = rs_Customers("Cust_Phone").Value
            txt_Address.Text = rs_Customers("Cust_Address").Value
            txt_Note.Text = rs_Customers("Cust_Note").Value
            C_name = rs_Customers("Cust_Name").Value
            C_id = rs_Customers("Cust_ID").Value
            TextBox2.Text = rs_Customers("Cust_Blance_Medin").Value
            Label7.Text = rs_Customers("Cust_Blance_Type").Value
            rs_Customers.Close()
            ComboBox1.Text = ""
            Label7.Visible = True
            ComboBox2.Visible = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select SUM(Stor_Quntity_No) as c_No From Stor Where Stor_DelFlag='" & s & "' And Cust_ID=" & C_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                TextBox1.Text = 0
            Else
                TextBox1.Text = rs_StoreVend("c_No").Value
            End If
            rs_StoreVend.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم العميل", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            TextBox2.Text = 0
        End If
        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Address.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر نوع الرصيد", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If
        'Dim x As Integer = getUser_ID() + 1
        Try
            Dim s As Boolean
            s = False
            Dim uk As Integer
            uk = 2
            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Select * From Customers Where Cust_Name='" & txt_Name.Text & "' And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Customers.EOF Or rs_Customers.BOF Then
                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From Customers Where Cust_ID= (SELECT MAX(Cust_ID)  FROM Customers)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("Cust_ID").Value
                'u_name1 = u_name + 1

                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Customers", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Customers.AddNew()
                'rs_Customers("Cust_ID").Value = u_name1
                rs_Customers("Cust_Name").Value = txt_Name.Text
                rs_Customers("Cust_Phone").Value = txt_Phone.Text
                rs_Customers("Cust_Address").Value = txt_Address.Text
                rs_Customers("Cust_Note").Value = txt_Note.Text
                rs_Customers("Cust_DelFlag").Value = s
                rs_Customers("User_ID").Value = u_Id
                rs_Customers("Cust_Blance_Medin").Value = TextBox2.Text
                rs_Customers("Cust_Blance_Type").Value = ComboBox2.Text
                rs_Customers("Cust_visable").Value = s
                rs_Customers("User_ID_edit").Value = uk
                rs_Customers("User_ID_delete").Value = uk
                If ComboBox2.Text = "عليه رصيد" Then
                    rs_Customers("Cust_Blance_d").Value = "defulit"
                End If
                If ComboBox2.Text = "ليه رصيد" Then
                    rs_Customers("Cust_Blance_d").Value = "notdefulit"
                End If
                rs_Customers.Update()
                MsgBox("تم حفظ العميل بنجاح", MsgBoxStyle.Information, "حفظ بيانات العميل")
                btn_All.Enabled = True
                btn_etite.Enabled = True
                btn_delete.Enabled = True
                btn_add.Enabled = True
                ComboBox1.Enabled = True
                txt_Name.Text = ""
                txt_Phone.Text = ""
                TextBox2.Text = ""
                txt_Address.Text = ""
                txt_Note.Text = ""
                btn_save.Enabled = False
                TextBox2.Enabled = False
                rs_Customers.Close()
                ComboBox2.Visible = False
                Label7.Visible = False
            Else
                MsgBox("أسم العميل موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                txt_Name.Text = ""
                txt_Name.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Me.Dispose()
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click

        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم العميل", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Address.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Select * From Customers Where Cust_Name='" & C_name & "' And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            rs_Customers("Cust_Name").Value = txt_Name.Text
            rs_Customers("Cust_Phone").Value = txt_Phone.Text
            rs_Customers("Cust_Address").Value = txt_Address.Text
            rs_Customers("Cust_Note").Value = txt_Note.Text
            rs_Customers("User_ID_edit").Value = u_Id
            'rs_Customers("Cust_Blance_Medin").Value = TextBox2.Text
            rs_Customers.Update()
            MsgBox("تم تعديل العميل بنجاح", MsgBoxStyle.Information, "تعديل بيانات المورد")
            txt_Name.Text = ""
            txt_Phone.Text = ""
            txt_Address.Text = ""
            txt_Note.Text = ""
            TextBox2.Text = ""
            Label7.Visible = False
            ComboBox2.Visible = False
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            rs_Customers.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم العميل", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        Try

            Dim s As Boolean
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Stor Where Cust_ID='" & C_id & "' And Stor_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                Dim g As String
                g = MsgBox("هل تريد حذف هذا العميل ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then
                    
                    If rs_Customers.State = 1 Then rs_Customers.Close()
                    rs_Customers.Open("Select * From Customers Where Cust_ID='" & C_id & "' And Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    s = True
                    rs_Customers("Cust_DelFlag").Value = s
                    rs_Customers("User_ID_delete").Value = u_Id
                    rs_Customers.Update()
                    MsgBox("تم حذف العميل بنجاح", MsgBoxStyle.Information, "حذف بيانات العميل")
                    txt_Name.Text = ""
                    txt_Phone.Text = ""
                    txt_Address.Text = ""
                    txt_Note.Text = ""
                    TextBox2.Text = ""
                    ComboBox1.Items.Clear()
                    ComboBox1.Text = ""
                    rs_Customers.Close()
                    Label7.Visible = False
                    ComboBox2.Visible = False
                End If
            Else
                MsgBox("لا يمكن حذف هذا العميل لوجود أذونات صادر فى المخزن", MsgBoxStyle.Information, "تحذير")

            End If


            
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            Dim t As Integer
            t = 1000
            If Val(TextBox1.Text) > t Then
                Label15.Text = "طـن"
            ElseIf Val(TextBox1.Text) <= t Then
                Label15.Text = "كيلـو"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        Frm_ShowCust.ShowDialog()
    End Sub

    Private Sub txt_Phone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Phone.KeyPress

        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If

    End Sub

    Private Sub Frm_Customers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_cust_add").Value
            auth_show = rs_auth("User_cust_show").Value
            auth_search = rs_auth("User_cust_search").Value
            auth_edit = rs_auth("User_cust_edit").Value
            auth_delete = rs_auth("User_cust_delete").Value
            If auth_add = ut Then
                btn_add.Visible = False
                btn_save.Visible = False
            Else
                btn_add.Visible = True
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
            If u_Id = 1 Then
                btn_add.Visible = True
                btn_save.Visible = True
                btn_All.Visible = True
                GroupBox3.Visible = True
                btn_etite.Visible = True
                btn_delete.Visible = True
            End If
            ComboBox2.Visible = False
            Label7.Visible = False
            txt_Name.Text = ""
            txt_Phone.Text = ""
            txt_Address.Text = ""
            txt_Note.Text = ""
            TextBox2.Text = ""
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
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

   
    Private Sub txt_Phone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Phone.TextChanged

    End Sub
End Class