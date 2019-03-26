Imports ADODB
Public Class Frm_Vendors
    Public v_name As String
    Public v_id As Integer


    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        ComboBox2.Visible = True
        Label7.Visible = False
        txt_Name.Text = ""
        TextBox2.Text = ""
        txt_Phone.Text = "014"
        txt_Address.Text = "لا يوجد عنوان"
        txt_Note.Text = ""
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = "ليه رصيد"
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
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            Dim s As Boolean
            s = False
            rs_Vendors.Open("Select * From Vendores Where Vend_DelFlag='" & s & "'  And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_Vendors.EOF
                ComboBox1.Items.Add(rs_Vendors("Vend_Name").Value)
                rs_Vendors.MoveNext()
            Loop
            rs_Vendors.Close()
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
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & u_name & "' And Vend_DelFlag='" & s & "' And Vend_visable='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            txt_Name.Text = rs_Vendors("Vend_Name").Value
            txt_Phone.Text = rs_Vendors("Vend_Phone").Value
            txt_Address.Text = rs_Vendors("Vend_Address").Value
            txt_Note.Text = rs_Vendors("Vend_Note").Value
            v_name = rs_Vendors("Vend_Name").Value
            v_id = rs_Vendors("Vend_ID").Value
            TextBox2.Text = rs_Vendors("Vend_Blance_Dein").Value
            Label7.Text = rs_Vendors("Vend_Blance_Type").Value
            rs_Vendors.Close()
            Label7.Visible = True
            ComboBox2.Visible = False
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select SUM(Stor_Quntity_No) as c_No From Stor Where Stor_DelFlag='" & s & "' And Vend_ID=" & v_id & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
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
            MsgBox("من فضلك أدخل أسم المورد", MsgBoxStyle.Information, "تنبيه")
            txt_Name.Select()
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            TextBox2.Text = 0
        End If
        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            txt_Phone.Select()
            Exit Sub
        End If

        If txt_Address.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            txt_Address.Select()
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
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & txt_Name.Text & "' And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From Vendores Where Vend_ID= (SELECT MAX(Vend_ID)  FROM Vendores)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("Vend_ID").Value
                'u_name1 = u_name + 1

                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Vendores", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Vendors.AddNew()
                'rs_Vendors("Vend_ID").Value = u_name1
                rs_Vendors("Vend_Name").Value = txt_Name.Text
                rs_Vendors("Vend_Phone").Value = txt_Phone.Text
                rs_Vendors("Vend_Address").Value = txt_Address.Text
                rs_Vendors("Vend_Note").Value = txt_Note.Text
                rs_Vendors("Vend_DelFlag").Value = s
                rs_Vendors("User_ID").Value = u_Id
                rs_Vendors("User_ID_edit").Value = uk
                rs_Vendors("User_ID_delete").Value = uk
                rs_Vendors("Vend_Blance_Dein").Value = TextBox2.Text
                rs_Vendors("Vend_Blance_Type").Value = ComboBox2.Text
                rs_Vendors("Vend_visable").Value = s
                If ComboBox2.Text = "ليه رصيد" Then
                    rs_Vendors("Vend_Blance_d").Value = "defulit"
                End If
                If ComboBox2.Text = "عليه رصيد" Then
                    rs_Vendors("Vend_Blance_d").Value = "notdefulit"
                End If
                rs_Vendors.Update()
                MsgBox("تم حفظ المورد بنجاح", MsgBoxStyle.Information, "حفظ بيانات المورد")
                btn_All.Enabled = True
                btn_etite.Enabled = True
                btn_delete.Enabled = True
                btn_add.Enabled = True
                ComboBox1.Enabled = True
                TextBox2.Enabled = False
                txt_Name.Text = ""
                TextBox2.Text = ""
                txt_Phone.Text = ""
                txt_Address.Text = ""
                txt_Note.Text = ""
                btn_save.Enabled = False
                ComboBox2.Visible = False
                Label7.Visible = False
                rs_Vendors.Close()
            Else
                MsgBox("أسم المورد موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                txt_Name.Text = ""
                txt_Name.Select()
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

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم المورد", MsgBoxStyle.Information, "تنبيه")
            txt_Name.Select()
            Exit Sub
        End If

        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            txt_Phone.Select()
            Exit Sub
        End If

        If txt_Address.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            txt_Address.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Vendores Where Vend_Name='" & v_name & "' And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            rs_Vendors("Vend_Name").Value = txt_Name.Text
            rs_Vendors("Vend_Phone").Value = txt_Phone.Text
            rs_Vendors("Vend_Address").Value = txt_Address.Text
            rs_Vendors("Vend_Note").Value = txt_Note.Text
            rs_Vendors("User_ID_edit").Value = u_Id
            'rs_Vendors("Vend_Blance_Dein").Value = TextBox2.Text
            rs_Vendors.Update()
            MsgBox("تم تعديل المورد بنجاح", MsgBoxStyle.Information, "تعديل بيانات المورد")
            txt_Name.Text = ""
            txt_Phone.Text = ""
            txt_Address.Text = ""
            TextBox2.Text = ""
            txt_Note.Text = ""

            ComboBox2.Visible = False
            Label7.Visible = False
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            rs_Vendors.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم المورد", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Stor Where Vend_ID='" & v_id & "' And Stor_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                Dim g As String
                g = MsgBox("هل تريد حذف هذا المورد ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then

                    If rs_Vendors.State = 1 Then rs_Vendors.Close()
                    rs_Vendors.Open("Select * From Vendores Where Vend_ID='" & v_id & "' And Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    s = True
                    rs_Vendors("Vend_DelFlag").Value = s
                    rs_Vendors("User_ID_delete").Value = u_Id
                    rs_Vendors.Update()
                    MsgBox("تم حذف المورد بنجاح", MsgBoxStyle.Information, "حذف بيانات المورد")
                    txt_Name.Text = ""
                    txt_Phone.Text = ""
                    txt_Address.Text = ""
                    txt_Note.Text = ""
                    TextBox2.Text = ""
                    ComboBox1.Items.Clear()
                    ComboBox1.Text = ""
                    rs_Vendors.Close()
                    ComboBox2.Visible = False
                    Label7.Visible = False
                End If
            Else
                MsgBox("لا يمكن حذف هذا المورد لوجود أذونات وارد فى المخزن", MsgBoxStyle.Information, "تحذير")
               
            End If
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
    Private Sub Frm_Vendors_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Try
            Me.Dispose()
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
        Frm_ShowVend.ShowDialog()
    End Sub

    Private Sub txt_Phone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Phone.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub



    Private Sub Frm_Vendors_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_vend_add").Value
            auth_show = rs_auth("User_vend_show").Value
            auth_search = rs_auth("User_vend_search").Value
            auth_edit = rs_auth("User_vend_edit").Value
            auth_delete = rs_auth("User_vend_delete").Value
            'auth_report = rs_auth("User_vend_report").Value
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
            TextBox2.Text = ""
            txt_Phone.Text = ""
            txt_Address.Text = ""
            txt_Note.Text = ""
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class