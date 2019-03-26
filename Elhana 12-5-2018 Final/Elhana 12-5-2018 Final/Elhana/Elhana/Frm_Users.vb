Imports ADODB
Public Class Frm_Users
    Public c_username As String
    Public jjj As Integer
    Public Function getUser_ID()

        If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
        Dim s As Boolean
        s = False
        rs_SofUserName.Open("Select Max(User_ID) as nu_ID From Users ", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Dim g As Integer
        g = rs_SofUserName("nu_ID").Value
        rs_SofUserName.Close()
        Return g


    End Function
    Public Sub ch_cleare()
        ch_broke.Checked = False
        ch_stor.Checked = False
        ch_blance.Checked = False
        ch_items.Checked = False
        ch_addmonth.Checked = False
        ch_addmonth.Checked = False
        ch_emp_add.Checked = False
        ch_emp_Show.Checked = False
        ch_emp_Search.Checked = False
        ch_emp_edit.Checked = False
        ch_emp_delete.Checked = False
        ch_emp_report.Checked = False
        ch_empatt_add.Checked = False
        ch_empatt_show.Checked = False
        ch_empatt_search.Checked = False
        ch_empatt_edit.Checked = False
        ch_empatt_delete.Checked = False
        ch_empatt_report.Checked = False
        ch_vend_add.Checked = False
        ch_vend_show.Checked = False
        ch_vend_search.Checked = False
        ch_vend_edit.Checked = False
        ch_vend_delete.Checked = False
        ch_vend_report.Checked = False
        ch_cust_add.Checked = False
        ch_cust_show.Checked = False
        ch_cust_search.Checked = False
        ch_cust_edit.Checked = False
        ch_cust_delete.Checked = False
        ch_cust_report.Checked = False
        ch_storvend_add.Checked = False
        ch_storvend_show.Checked = False
        ch_storvend_search.Checked = False
        ch_storvend_edit.Checked = False
        ch_storvend_delete.Checked = False
        ch_storvend_report.Checked = False
        ch_storcust_add.Checked = False
        ch_storcust_show.Checked = False
        ch_storcust_search.Checked = False
        ch_storcust_edit.Checked = False
        ch_storcust_delete.Checked = False
        ch_storcust_report.Checked = False
        ch_moneyimp_add.Checked = False
        ch_moneyimp_show.Checked = False
        ch_moneyimp_search.Checked = False
        ch_moneyimp_edit.Checked = False
        ch_moneyimp_delete.Checked = False
        ch_moneyimp_report.Checked = False
        ch_moneyexp_add.Checked = False
        ch_moneyexp_show.Checked = False
        ch_moneyexp_search.Checked = False
        ch_moneyexp_edit.Checked = False
        ch_moneyexp_delete.Checked = False
        ch_moneyexp_report.Checked = False
        ch_user_add.Checked = False
        ch_user_show.Checked = False
        ch_user_search.Checked = False
        ch_user_edit.Checked = False
    End Sub
    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Try
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            Dim uuu As Integer
            uuu = 1
            If u_Id = uuu Then

                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                Dim s As Boolean
                s = False
                rs_SofUserName.Open("Select * From Users Where User_Flag='" & s & "' And User_visbale='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                ComboBox1.Items.Clear()
                Do While Not rs_SofUserName.EOF
                    ComboBox1.Items.Add(rs_SofUserName("User_Name").Value)
                    rs_SofUserName.MoveNext()
                Loop
            Else
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                Dim s As Boolean
                s = False
                rs_SofUserName.Open("Select * From Users Where User_Flag='" & s & "' And User_ID > " & uuu & " And User_visbale='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                ComboBox1.Items.Clear()
                Do While Not rs_SofUserName.EOF
                    ComboBox1.Items.Add(rs_SofUserName("User_Name").Value)
                    rs_SofUserName.MoveNext()
                Loop
            End If

           
            rs_SofUserName.Close()
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
            Dim uuu As Integer
            uuu = 1
            If u_Id = uuu Then
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                rs_SofUserName.Open("Select * From Users Where User_Name='" & u_name & "' And User_Flag='" & s & "' And User_visbale='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                jjj = rs_SofUserName("User_ID").Value
                txt_Name.Text = rs_SofUserName("User_Name").Value
                txt_Pass.Text = rs_SofUserName("User_Password").Value

                If jjj = 1 Then
                    txt_Name.Enabled = False
                    CheckBox1.Enabled = False
                    ch_broke.Enabled = False
                    GroupBox4.Enabled = False
                    GroupBox5.Enabled = False
                    GroupBox6.Enabled = False
                    GroupBox7.Enabled = False
                    GroupBox8.Enabled = False
                    GroupBox9.Enabled = False
                    GroupBox10.Enabled = False
                    GroupBox11.Enabled = False
                    GroupBox12.Enabled = False
                    GroupBox13.Enabled = False
                    rd_Admin.Enabled = False
                    rd_Employee.Enabled = False
                    rd_manager.Enabled = False
                Else
                    txt_Name.Enabled = True
                    CheckBox1.Enabled = True
                    ch_broke.Enabled = True
                    GroupBox4.Enabled = True
                    GroupBox5.Enabled = True
                    GroupBox6.Enabled = True
                    GroupBox7.Enabled = True
                    GroupBox8.Enabled = True
                    GroupBox9.Enabled = True
                    GroupBox10.Enabled = True
                    GroupBox11.Enabled = True
                    GroupBox12.Enabled = True
                    GroupBox13.Enabled = True
                    rd_Admin.Enabled = True
                    rd_Employee.Enabled = True
                    rd_manager.Enabled = True
                    CheckBox1.Checked = rs_SofUserName("User_State").Value
                    ch_broke.Checked = rs_SofUserName("User_State_brok").Value
                    c_username = rs_SofUserName("User_Name").Value

                    If rs_SofUserName("User_type").Value = "admin" Then
                        rd_Admin.Checked = True
                    End If
                    If rs_SofUserName("User_type").Value = "manager" Then
                        rd_manager.Checked = True
                    End If
                    If rs_SofUserName("User_type").Value = "Employee" Then
                        rd_Employee.Checked = True
                    End If
                    ch_stor.Checked = rs_SofUserName("User_stor").Value
                    ch_blance.Checked = rs_SofUserName("User_blance").Value
                    ch_items.Checked = rs_SofUserName("User_items").Value
                    ch_addmonth.Checked = rs_SofUserName("User_addmonth").Value
                    ch_emp_add.Checked = rs_SofUserName("User_emp_add").Value
                    ch_emp_Show.Checked = rs_SofUserName("User_emp_show").Value
                    ch_emp_Search.Checked = rs_SofUserName("User_emp_search").Value
                    ch_emp_edit.Checked = rs_SofUserName("User_emp_edit").Value
                    ch_emp_delete.Checked = rs_SofUserName("User_emp_delete").Value
                    ch_emp_report.Checked = rs_SofUserName("User_emp_report").Value
                    ch_empatt_add.Checked = rs_SofUserName("User_empatt_add").Value
                    ch_empatt_show.Checked = rs_SofUserName("User_empatt_show").Value
                    ch_empatt_search.Checked = rs_SofUserName("User_empatt_search").Value
                    ch_empatt_edit.Checked = rs_SofUserName("User_empatt_edit").Value
                    ch_empatt_delete.Checked = rs_SofUserName("User_empatt_delete").Value
                    ch_empatt_report.Checked = rs_SofUserName("User_empatt_report").Value
                    ch_vend_add.Checked = rs_SofUserName("User_vend_add").Value
                    ch_vend_show.Checked = rs_SofUserName("User_vend_show").Value
                    ch_vend_search.Checked = rs_SofUserName("User_vend_search").Value
                    ch_vend_edit.Checked = rs_SofUserName("User_vend_edit").Value
                    ch_vend_delete.Checked = rs_SofUserName("User_vend_delete").Value
                    ch_vend_report.Checked = rs_SofUserName("User_vend_report").Value
                    ch_cust_add.Checked = rs_SofUserName("User_cust_add").Value
                    ch_cust_show.Checked = rs_SofUserName("User_cust_show").Value
                    ch_cust_search.Checked = rs_SofUserName("User_cust_search").Value
                    ch_cust_edit.Checked = rs_SofUserName("User_cust_edit").Value
                    ch_cust_delete.Checked = rs_SofUserName("User_cust_delete").Value
                    ch_cust_report.Checked = rs_SofUserName("User_cust_report").Value
                    ch_storvend_add.Checked = rs_SofUserName("User_storvend_add").Value
                    ch_storvend_show.Checked = rs_SofUserName("User_storvend_show").Value
                    ch_storvend_search.Checked = rs_SofUserName("User_storvend_search").Value
                    ch_storvend_edit.Checked = rs_SofUserName("User_storvend_edit").Value
                    ch_storvend_delete.Checked = rs_SofUserName("User_storvend_delete").Value
                    ch_storvend_report.Checked = rs_SofUserName("User_storvend_report").Value
                    ch_storcust_add.Checked = rs_SofUserName("User_storcust_add").Value
                    ch_storcust_show.Checked = rs_SofUserName("User_storcust_show").Value
                    ch_storcust_search.Checked = rs_SofUserName("User_storcust_search").Value
                    ch_storcust_edit.Checked = rs_SofUserName("User_storcust_edit").Value
                    ch_storcust_delete.Checked = rs_SofUserName("User_storcust_delete").Value
                    ch_storcust_report.Checked = rs_SofUserName("User_storcust_reprt").Value
                    ch_moneyimp_add.Checked = rs_SofUserName("User_moneyimp_add").Value
                    ch_moneyimp_show.Checked = rs_SofUserName("User_moneyimp_show").Value
                    ch_moneyimp_search.Checked = rs_SofUserName("User_moneyimp_search").Value
                    ch_moneyimp_edit.Checked = rs_SofUserName("User_moneyimp_edit").Value
                    ch_moneyimp_delete.Checked = rs_SofUserName("User_moneyimp_delete").Value
                    ch_moneyimp_report.Checked = rs_SofUserName("User_moneyimp_repot").Value
                    ch_moneyexp_add.Checked = rs_SofUserName("User_moneyexp_add").Value
                    ch_moneyexp_show.Checked = rs_SofUserName("User_moneyexp_show").Value
                    ch_moneyexp_search.Checked = rs_SofUserName("User_moneyexp_search").Value
                    ch_moneyexp_edit.Checked = rs_SofUserName("User_moneyexp_edit").Value
                    ch_moneyexp_delete.Checked = rs_SofUserName("User_moneyexp_delete").Value
                    ch_moneyexp_report.Checked = rs_SofUserName("User_moneyexp_report").Value
                    ch_user_add.Checked = rs_SofUserName("User_user_add").Value
                    ch_user_show.Checked = rs_SofUserName("User_user_show").Value
                    ch_user_search.Checked = rs_SofUserName("User_user_search").Value
                    ch_user_edit.Checked = rs_SofUserName("User_user_edit").Value
                End If

            Else

                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                rs_SofUserName.Open("Select * From Users Where User_Name='" & u_name & "' And User_ID > " & uuu & " And User_Flag='" & s & "' And User_visbale='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                jjj = rs_SofUserName("User_ID").Value
                txt_Name.Text = rs_SofUserName("User_Name").Value
                txt_Pass.Text = rs_SofUserName("User_Password").Value

               
                    CheckBox1.Checked = rs_SofUserName("User_State").Value
                    ch_broke.Checked = rs_SofUserName("User_State_brok").Value
                    c_username = rs_SofUserName("User_Name").Value

                    If rs_SofUserName("User_type").Value = "admin" Then
                        rd_Admin.Checked = True
                    End If
                    If rs_SofUserName("User_type").Value = "manager" Then
                        rd_manager.Checked = True
                    End If
                    If rs_SofUserName("User_type").Value = "Employee" Then
                        rd_Employee.Checked = True
                    End If
                    ch_stor.Checked = rs_SofUserName("User_stor").Value
                    ch_blance.Checked = rs_SofUserName("User_blance").Value
                    ch_items.Checked = rs_SofUserName("User_items").Value
                    ch_addmonth.Checked = rs_SofUserName("User_addmonth").Value
                    ch_emp_add.Checked = rs_SofUserName("User_emp_add").Value
                    ch_emp_Show.Checked = rs_SofUserName("User_emp_show").Value
                    ch_emp_Search.Checked = rs_SofUserName("User_emp_search").Value
                    ch_emp_edit.Checked = rs_SofUserName("User_emp_edit").Value
                    ch_emp_delete.Checked = rs_SofUserName("User_emp_delete").Value
                    ch_emp_report.Checked = rs_SofUserName("User_emp_report").Value
                    ch_empatt_add.Checked = rs_SofUserName("User_empatt_add").Value
                    ch_empatt_show.Checked = rs_SofUserName("User_empatt_show").Value
                    ch_empatt_search.Checked = rs_SofUserName("User_empatt_search").Value
                    ch_empatt_edit.Checked = rs_SofUserName("User_empatt_edit").Value
                    ch_empatt_delete.Checked = rs_SofUserName("User_empatt_delete").Value
                    ch_empatt_report.Checked = rs_SofUserName("User_empatt_report").Value
                    ch_vend_add.Checked = rs_SofUserName("User_vend_add").Value
                    ch_vend_show.Checked = rs_SofUserName("User_vend_show").Value
                    ch_vend_search.Checked = rs_SofUserName("User_vend_search").Value
                    ch_vend_edit.Checked = rs_SofUserName("User_vend_edit").Value
                    ch_vend_delete.Checked = rs_SofUserName("User_vend_delete").Value
                    ch_vend_report.Checked = rs_SofUserName("User_vend_report").Value
                    ch_cust_add.Checked = rs_SofUserName("User_cust_add").Value
                    ch_cust_show.Checked = rs_SofUserName("User_cust_show").Value
                    ch_cust_search.Checked = rs_SofUserName("User_cust_search").Value
                    ch_cust_edit.Checked = rs_SofUserName("User_cust_edit").Value
                    ch_cust_delete.Checked = rs_SofUserName("User_cust_delete").Value
                    ch_cust_report.Checked = rs_SofUserName("User_cust_report").Value
                    ch_storvend_add.Checked = rs_SofUserName("User_storvend_add").Value
                    ch_storvend_show.Checked = rs_SofUserName("User_storvend_show").Value
                    ch_storvend_search.Checked = rs_SofUserName("User_storvend_search").Value
                    ch_storvend_edit.Checked = rs_SofUserName("User_storvend_edit").Value
                    ch_storvend_delete.Checked = rs_SofUserName("User_storvend_delete").Value
                    ch_storvend_report.Checked = rs_SofUserName("User_storvend_report").Value
                    ch_storcust_add.Checked = rs_SofUserName("User_storcust_add").Value
                    ch_storcust_show.Checked = rs_SofUserName("User_storcust_show").Value
                    ch_storcust_search.Checked = rs_SofUserName("User_storcust_search").Value
                    ch_storcust_edit.Checked = rs_SofUserName("User_storcust_edit").Value
                    ch_storcust_delete.Checked = rs_SofUserName("User_storcust_delete").Value
                    ch_storcust_report.Checked = rs_SofUserName("User_storcust_reprt").Value
                    ch_moneyimp_add.Checked = rs_SofUserName("User_moneyimp_add").Value
                    ch_moneyimp_show.Checked = rs_SofUserName("User_moneyimp_show").Value
                    ch_moneyimp_search.Checked = rs_SofUserName("User_moneyimp_search").Value
                    ch_moneyimp_edit.Checked = rs_SofUserName("User_moneyimp_edit").Value
                    ch_moneyimp_delete.Checked = rs_SofUserName("User_moneyimp_delete").Value
                    ch_moneyimp_report.Checked = rs_SofUserName("User_moneyimp_repot").Value
                    ch_moneyexp_add.Checked = rs_SofUserName("User_moneyexp_add").Value
                    ch_moneyexp_show.Checked = rs_SofUserName("User_moneyexp_show").Value
                    ch_moneyexp_search.Checked = rs_SofUserName("User_moneyexp_search").Value
                    ch_moneyexp_edit.Checked = rs_SofUserName("User_moneyexp_edit").Value
                    ch_moneyexp_delete.Checked = rs_SofUserName("User_moneyexp_delete").Value
                    ch_moneyexp_report.Checked = rs_SofUserName("User_moneyexp_report").Value
                    ch_user_add.Checked = rs_SofUserName("User_user_add").Value
                    ch_user_show.Checked = rs_SofUserName("User_user_show").Value
                    ch_user_search.Checked = rs_SofUserName("User_user_search").Value
                    ch_user_edit.Checked = rs_SofUserName("User_user_edit").Value


                End If


                rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        txt_Name.Text = ""
        txt_Pass.Text = ""
        ComboBox1.Items.Clear()
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        ComboBox1.Enabled = False
        btn_save.Enabled = True
        txt_Name.Enabled = True
        CheckBox1.Enabled = True
        ch_broke.Enabled = True
        GroupBox4.Enabled = True
        GroupBox5.Enabled = True
        GroupBox6.Enabled = True
        GroupBox7.Enabled = True
        GroupBox8.Enabled = True
        GroupBox9.Enabled = True
        GroupBox10.Enabled = True
        GroupBox11.Enabled = True
        GroupBox12.Enabled = True
        GroupBox13.Enabled = True
        rd_Admin.Enabled = True
        rd_Employee.Enabled = True
        rd_manager.Enabled = True
        ch_cleare()
        rd_Employee.Checked = True
        txt_Name.Select()
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Pass.Text = "" Then
            MsgBox("من فضلك أدخل كلمة المرور", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        Try
            Dim x As Integer = getUser_ID() + 1
            Dim s As Boolean
            s = False
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.Open("Select * From Users Where User_Name='" & txt_Name.Text & "' And User_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_SofUserName.EOF Or rs_SofUserName.BOF Then
                
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                rs_SofUserName.Open("Users", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_SofUserName.AddNew()
                rs_SofUserName("User_Name").Value = txt_Name.Text
                rs_SofUserName("User_Password").Value = txt_Pass.Text
                rs_SofUserName("User_State").Value = CheckBox1.Checked
                rs_SofUserName("User_Flag").Value = s
                rs_SofUserName("User_State_brok").Value = ch_broke.Checked
                If rd_Admin.Checked = True Then
                    rs_SofUserName("User_type").Value = "admin"
                End If
                If rd_manager.Checked = True Then
                    rs_SofUserName("User_type").Value = "manager"
                End If
                If rd_Employee.Checked = True Then
                    rs_SofUserName("User_type").Value = "Employee"
                End If
                rs_SofUserName("User_stor").Value = ch_stor.Checked
                rs_SofUserName("User_blance").Value = ch_blance.Checked
                rs_SofUserName("User_items").Value = ch_items.Checked
                rs_SofUserName("User_addmonth").Value = ch_addmonth.Checked
                rs_SofUserName("User_emp_add").Value = ch_emp_add.Checked
                rs_SofUserName("User_emp_show").Value = ch_emp_Show.Checked
                rs_SofUserName("User_emp_search").Value = ch_emp_Search.Checked
                rs_SofUserName("User_emp_edit").Value = ch_emp_edit.Checked
                rs_SofUserName("User_emp_delete").Value = ch_emp_delete.Checked
                rs_SofUserName("User_emp_report").Value = ch_emp_report.Checked
                rs_SofUserName("User_empatt_add").Value = ch_empatt_add.Checked
                rs_SofUserName("User_empatt_show").Value = ch_empatt_show.Checked
                rs_SofUserName("User_empatt_search").Value = ch_empatt_search.Checked
                rs_SofUserName("User_empatt_edit").Value = ch_empatt_edit.Checked
                rs_SofUserName("User_empatt_delete").Value = ch_empatt_delete.Checked
                rs_SofUserName("User_empatt_report").Value = ch_empatt_report.Checked
                rs_SofUserName("User_vend_add").Value = ch_vend_add.Checked
                rs_SofUserName("User_vend_show").Value = ch_vend_show.Checked
                rs_SofUserName("User_vend_search").Value = ch_vend_search.Checked
                rs_SofUserName("User_vend_edit").Value = ch_vend_edit.Checked
                rs_SofUserName("User_vend_delete").Value = ch_vend_delete.Checked
                rs_SofUserName("User_vend_report").Value = ch_vend_report.Checked
                rs_SofUserName("User_cust_add").Value = ch_cust_add.Checked
                rs_SofUserName("User_cust_show").Value = ch_cust_show.Checked
                rs_SofUserName("User_cust_search").Value = ch_cust_search.Checked
                rs_SofUserName("User_cust_edit").Value = ch_cust_edit.Checked
                rs_SofUserName("User_cust_delete").Value = ch_cust_delete.Checked
                rs_SofUserName("User_cust_report").Value = ch_cust_report.Checked
                rs_SofUserName("User_storvend_add").Value = ch_storvend_add.Checked
                rs_SofUserName("User_storvend_show").Value = ch_storvend_show.Checked
                rs_SofUserName("User_storvend_search").Value = ch_storvend_search.Checked
                rs_SofUserName("User_storvend_edit").Value = ch_storvend_edit.Checked
                rs_SofUserName("User_storvend_delete").Value = ch_storvend_delete.Checked
                rs_SofUserName("User_storvend_report").Value = ch_storvend_report.Checked
                rs_SofUserName("User_storcust_add").Value = ch_storcust_add.Checked
                rs_SofUserName("User_storcust_show").Value = ch_storcust_show.Checked
                rs_SofUserName("User_storcust_search").Value = ch_storcust_search.Checked
                rs_SofUserName("User_storcust_edit").Value = ch_storcust_edit.Checked
                rs_SofUserName("User_storcust_delete").Value = ch_storcust_delete.Checked
                rs_SofUserName("User_storcust_reprt").Value = ch_storcust_report.Checked
                rs_SofUserName("User_moneyimp_add").Value = ch_moneyimp_add.Checked
                rs_SofUserName("User_moneyimp_show").Value = ch_moneyimp_show.Checked
                rs_SofUserName("User_moneyimp_search").Value = ch_moneyimp_search.Checked
                rs_SofUserName("User_moneyimp_edit").Value = ch_moneyimp_edit.Checked
                rs_SofUserName("User_moneyimp_delete").Value = ch_moneyimp_delete.Checked
                rs_SofUserName("User_moneyimp_repot").Value = ch_moneyimp_report.Checked
                rs_SofUserName("User_moneyexp_add").Value = ch_moneyexp_add.Checked
                rs_SofUserName("User_moneyexp_show").Value = ch_moneyexp_show.Checked
                rs_SofUserName("User_moneyexp_search").Value = ch_moneyexp_search.Checked
                rs_SofUserName("User_moneyexp_edit").Value = ch_moneyexp_edit.Checked
                rs_SofUserName("User_moneyexp_delete").Value = ch_moneyexp_delete.Checked
                rs_SofUserName("User_moneyexp_report").Value = ch_moneyexp_report.Checked
                rs_SofUserName("User_user_add").Value = ch_user_add.Checked
                rs_SofUserName("User_user_show").Value = ch_user_show.Checked
                rs_SofUserName("User_user_search").Value = ch_user_search.Checked
                rs_SofUserName("User_user_edit").Value = ch_user_edit.Checked
                rs_SofUserName("User_visbale").Value = False
                rs_SofUserName.Update()
                MsgBox("تم حفظ المستخدم بنجاح", MsgBoxStyle.Information, "حفظ بيانات المستخدم")
                btn_All.Enabled = True
                btn_etite.Enabled = True
                btn_delete.Enabled = True
                btn_add.Enabled = True
                ComboBox1.Enabled = True
                txt_Name.Text = ""
                txt_Pass.Text = ""
                CheckBox1.Checked = False
                btn_save.Enabled = False
                ch_cleare()
                rs_SofUserName.Close()
            Else
                MsgBox("أسم المستخدم موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                txt_Name.Text = ""
                txt_Name.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Pass.Text = "" Then
            MsgBox("من فضلك أدخل كلمة المرور", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.Open("Select * From Users Where User_ID=" & jjj & " And User_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_SofUserName.EOF Or rs_SofUserName.BOF Then
                MsgBox("أسم المستخدم غير موجود من فضلك أختر مستخدم أخر", MsgBoxStyle.Information, "تحذير")
                ComboBox1.Text = ""
                ComboBox1.Select()
            Else

                If jjj = 1 Then
                    rs_SofUserName("User_Password").Value = txt_Pass.Text
                    rs_SofUserName.Update()
                    MsgBox("تم تعديل المستخدم بنجاح", MsgBoxStyle.Information, "تعديل بيانات المستخدم")
                    txt_Name.Text = ""
                    txt_Pass.Text = ""
                    CheckBox1.Checked = False
                    ComboBox1.Items.Clear()
                    ComboBox1.Text = ""
                Else



                    rs_SofUserName("User_Name").Value = txt_Name.Text
                    rs_SofUserName("User_Password").Value = txt_Pass.Text
                    rs_SofUserName("User_State").Value = CheckBox1.Checked
                    rs_SofUserName("User_State_brok").Value = ch_broke.Checked
                    If rd_Admin.Checked = True Then
                        rs_SofUserName("User_type").Value = "admin"
                    End If
                    If rd_manager.Checked = True Then
                        rs_SofUserName("User_type").Value = "manager"
                    End If
                    If rd_Employee.Checked = True Then
                        rs_SofUserName("User_type").Value = "Employee"
                    End If
                    rs_SofUserName("User_stor").Value = ch_stor.Checked
                    rs_SofUserName("User_blance").Value = ch_blance.Checked
                    rs_SofUserName("User_items").Value = ch_items.Checked
                    rs_SofUserName("User_addmonth").Value = ch_addmonth.Checked
                    rs_SofUserName("User_emp_add").Value = ch_emp_add.Checked
                    rs_SofUserName("User_emp_show").Value = ch_emp_Show.Checked
                    rs_SofUserName("User_emp_search").Value = ch_emp_Search.Checked
                    rs_SofUserName("User_emp_edit").Value = ch_emp_edit.Checked
                    rs_SofUserName("User_emp_delete").Value = ch_emp_delete.Checked
                    rs_SofUserName("User_emp_report").Value = ch_emp_report.Checked
                    rs_SofUserName("User_empatt_add").Value = ch_empatt_add.Checked
                    rs_SofUserName("User_empatt_show").Value = ch_empatt_show.Checked
                    rs_SofUserName("User_empatt_search").Value = ch_empatt_search.Checked
                    rs_SofUserName("User_empatt_edit").Value = ch_empatt_edit.Checked
                    rs_SofUserName("User_empatt_delete").Value = ch_empatt_delete.Checked
                    rs_SofUserName("User_empatt_report").Value = ch_empatt_report.Checked
                    rs_SofUserName("User_vend_add").Value = ch_vend_add.Checked
                    rs_SofUserName("User_vend_show").Value = ch_vend_show.Checked
                    rs_SofUserName("User_vend_search").Value = ch_vend_search.Checked
                    rs_SofUserName("User_vend_edit").Value = ch_vend_edit.Checked
                    rs_SofUserName("User_vend_delete").Value = ch_vend_delete.Checked
                    rs_SofUserName("User_vend_report").Value = ch_vend_report.Checked
                    rs_SofUserName("User_cust_add").Value = ch_cust_add.Checked
                    rs_SofUserName("User_cust_show").Value = ch_cust_show.Checked
                    rs_SofUserName("User_cust_search").Value = ch_cust_search.Checked
                    rs_SofUserName("User_cust_edit").Value = ch_cust_edit.Checked
                    rs_SofUserName("User_cust_delete").Value = ch_cust_delete.Checked
                    rs_SofUserName("User_cust_report").Value = ch_cust_report.Checked
                    rs_SofUserName("User_storvend_add").Value = ch_storvend_add.Checked
                    rs_SofUserName("User_storvend_show").Value = ch_storvend_show.Checked
                    rs_SofUserName("User_storvend_search").Value = ch_storvend_search.Checked
                    rs_SofUserName("User_storvend_edit").Value = ch_storvend_edit.Checked
                    rs_SofUserName("User_storvend_delete").Value = ch_storvend_delete.Checked
                    rs_SofUserName("User_storvend_report").Value = ch_storvend_report.Checked
                    rs_SofUserName("User_storcust_add").Value = ch_storcust_add.Checked
                    rs_SofUserName("User_storcust_show").Value = ch_storcust_show.Checked
                    rs_SofUserName("User_storcust_search").Value = ch_storcust_search.Checked
                    rs_SofUserName("User_storcust_edit").Value = ch_storcust_edit.Checked
                    rs_SofUserName("User_storcust_delete").Value = ch_storcust_delete.Checked
                    rs_SofUserName("User_storcust_reprt").Value = ch_storcust_report.Checked
                    rs_SofUserName("User_moneyimp_add").Value = ch_moneyimp_add.Checked
                    rs_SofUserName("User_moneyimp_show").Value = ch_moneyimp_show.Checked
                    rs_SofUserName("User_moneyimp_search").Value = ch_moneyimp_search.Checked
                    rs_SofUserName("User_moneyimp_edit").Value = ch_moneyimp_edit.Checked
                    rs_SofUserName("User_moneyimp_delete").Value = ch_moneyimp_delete.Checked
                    rs_SofUserName("User_moneyimp_repot").Value = ch_moneyimp_report.Checked
                    rs_SofUserName("User_moneyexp_add").Value = ch_moneyexp_add.Checked
                    rs_SofUserName("User_moneyexp_show").Value = ch_moneyexp_show.Checked
                    rs_SofUserName("User_moneyexp_search").Value = ch_moneyexp_search.Checked
                    rs_SofUserName("User_moneyexp_edit").Value = ch_moneyexp_edit.Checked
                    rs_SofUserName("User_moneyexp_delete").Value = ch_moneyexp_delete.Checked
                    rs_SofUserName("User_moneyexp_report").Value = ch_moneyexp_report.Checked
                    rs_SofUserName("User_user_add").Value = ch_user_add.Checked
                    rs_SofUserName("User_user_show").Value = ch_user_show.Checked
                    rs_SofUserName("User_user_search").Value = ch_user_search.Checked
                    rs_SofUserName("User_user_edit").Value = ch_user_edit.Checked
                    rs_SofUserName.Update()
                    MsgBox("تم تعديل المستخدم بنجاح", MsgBoxStyle.Information, "تعديل بيانات المستخدم")
                    txt_Name.Text = ""
                    txt_Pass.Text = ""
                    CheckBox1.Checked = False
                    ComboBox1.Items.Clear()
                    ComboBox1.Text = ""
                End If
                rs_SofUserName.Close()
                txt_Name.Enabled = True
                CheckBox1.Enabled = True
                ch_broke.Enabled = True
                GroupBox4.Enabled = True
                GroupBox5.Enabled = True
                GroupBox6.Enabled = True
                GroupBox7.Enabled = True
                GroupBox8.Enabled = True
                GroupBox9.Enabled = True
                GroupBox10.Enabled = True
                GroupBox11.Enabled = True
                GroupBox12.Enabled = True
                GroupBox13.Enabled = True
                rd_Admin.Enabled = True
                rd_Employee.Enabled = True
                rd_manager.Enabled = True
                ch_cleare()
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If

        Try
            Dim g As String
            g = MsgBox("هل تريد حذف هذا المستخدم ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

            If g = vbYes Then

                Dim s As Boolean
                s = False
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                rs_SofUserName.Open("Select * From Users Where User_Name='" & c_username & "' And User_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                s = True
                rs_SofUserName("User_Flag").Value = s
                rs_SofUserName.Update()
                MsgBox("تم حذف المستخدم بنجاح", MsgBoxStyle.Information, "حذف بيانات المستخدم")
                txt_Name.Text = ""
                txt_Pass.Text = ""
                ComboBox1.Text = ""
                CheckBox1.Checked = False
                ComboBox1.Items.Clear()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        frm_ShowAllUser.ShowDialog()
    End Sub

    Private Sub Frm_Users_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_user_add").Value
            auth_show = rs_auth("User_user_show").Value
            auth_search = rs_auth("User_user_search").Value
            auth_edit = rs_auth("User_user_edit").Value
            
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
            
            If u_Id = 1 Then
                btn_add.Visible = True
                btn_save.Visible = True
                btn_All.Visible = True
                GroupBox3.Visible = True
                btn_etite.Visible = True

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ch_emp_add_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_add.CheckedChanged
        ch_empatt_add.Checked = ch_emp_add.Checked
        ch_vend_add.Checked = ch_emp_add.Checked
        ch_cust_add.Checked = ch_emp_add.Checked
        ch_storvend_add.Checked = ch_emp_add.Checked
        ch_storcust_add.Checked = ch_emp_add.Checked
        ch_moneyimp_add.Checked = ch_emp_add.Checked
        ch_moneyexp_add.Checked = ch_emp_add.Checked
    End Sub

    Private Sub ch_emp_Show_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_Show.CheckedChanged
        ch_empatt_show.Checked = ch_emp_Show.Checked
        ch_vend_show.Checked = ch_emp_Show.Checked
        ch_cust_show.Checked = ch_emp_Show.Checked
        ch_storvend_show.Checked = ch_emp_Show.Checked
        ch_storcust_show.Checked = ch_emp_Show.Checked
        ch_moneyimp_show.Checked = ch_emp_Show.Checked
        ch_moneyexp_show.Checked = ch_emp_Show.Checked
    End Sub

    Private Sub ch_emp_Search_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_Search.CheckedChanged
        ch_empatt_search.Checked = ch_emp_Search.Checked
        ch_vend_search.Checked = ch_emp_Search.Checked
        ch_cust_search.Checked = ch_emp_Search.Checked
        ch_storvend_search.Checked = ch_emp_Search.Checked
        ch_storcust_search.Checked = ch_emp_Search.Checked
        ch_moneyimp_search.Checked = ch_emp_Search.Checked
        ch_moneyexp_search.Checked = ch_emp_Search.Checked
    End Sub

    Private Sub ch_emp_edit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_edit.CheckedChanged
        ch_empatt_edit.Checked = ch_emp_edit.Checked
        ch_vend_edit.Checked = ch_emp_edit.Checked
        ch_cust_edit.Checked = ch_emp_edit.Checked
        ch_storvend_edit.Checked = ch_emp_edit.Checked
        ch_storcust_edit.Checked = ch_emp_edit.Checked
        ch_moneyimp_edit.Checked = ch_emp_edit.Checked
        ch_moneyexp_edit.Checked = ch_emp_edit.Checked
    End Sub

    Private Sub ch_emp_delete_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_delete.CheckedChanged
        ch_empatt_delete.Checked = ch_emp_delete.Checked
        ch_vend_delete.Checked = ch_emp_delete.Checked
        ch_cust_delete.Checked = ch_emp_delete.Checked
        ch_storvend_delete.Checked = ch_emp_delete.Checked
        ch_storcust_delete.Checked = ch_emp_delete.Checked
        ch_moneyimp_delete.Checked = ch_emp_delete.Checked
        ch_moneyexp_delete.Checked = ch_emp_delete.Checked
    End Sub

    Private Sub ch_emp_report_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_emp_report.CheckedChanged
        ch_empatt_report.Checked = ch_emp_report.Checked
        ch_vend_report.Checked = ch_emp_report.Checked
        ch_cust_report.Checked = ch_emp_report.Checked
        ch_storvend_report.Checked = ch_emp_report.Checked
        ch_storcust_report.Checked = ch_emp_report.Checked
        ch_moneyimp_report.Checked = ch_emp_report.Checked
        ch_moneyexp_report.Checked = ch_emp_report.Checked
    End Sub

    Private Sub rd_Admin_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rd_Admin.CheckedChanged
        ch_stor.Checked = True
        ch_blance.Checked = True
        ch_items.Checked = True
        ch_addmonth.Checked = True
        ch_emp_add.Checked = True
        ch_emp_Show.Checked = True
        ch_emp_Search.Checked = True
        ch_emp_edit.Checked = True
        ch_emp_delete.Checked = True
        ch_emp_report.Checked = True
        ch_empatt_add.Checked = True
        ch_empatt_show.Checked = True
        ch_empatt_search.Checked = True
        ch_empatt_edit.Checked = True
        ch_empatt_delete.Checked = True
        ch_empatt_report.Checked = True
        ch_vend_add.Checked = True
        ch_vend_show.Checked = True
        ch_vend_search.Checked = True
        ch_vend_edit.Checked = True
        ch_vend_delete.Checked = True
        ch_vend_report.Checked = True
        ch_cust_add.Checked = True
        ch_cust_show.Checked = True
        ch_cust_search.Checked = True
        ch_cust_edit.Checked = True
        ch_cust_delete.Checked = True
        ch_cust_report.Checked = True
        ch_storvend_add.Checked = True
        ch_storvend_show.Checked = True
        ch_storvend_search.Checked = True
        ch_storvend_edit.Checked = True
        ch_storvend_delete.Checked = True
        ch_storvend_report.Checked = True
        ch_storcust_add.Checked = True
        ch_storcust_show.Checked = True
        ch_storcust_search.Checked = True
        ch_storcust_edit.Checked = True
        ch_storcust_delete.Checked = True
        ch_storcust_report.Checked = True
        ch_moneyimp_add.Checked = True
        ch_moneyimp_show.Checked = True
        ch_moneyimp_search.Checked = True
        ch_moneyimp_edit.Checked = True
        ch_moneyimp_delete.Checked = True
        ch_moneyimp_report.Checked = True
        ch_moneyexp_add.Checked = True
        ch_moneyexp_show.Checked = True
        ch_moneyexp_search.Checked = True
        ch_moneyexp_edit.Checked = True
        ch_moneyexp_delete.Checked = True
        ch_moneyexp_report.Checked = True

        ch_user_add.Checked = True
        ch_user_show.Checked = True
        ch_user_search.Checked = True
        ch_user_edit.Checked = True
       
    End Sub

    Private Sub rd_manager_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rd_manager.CheckedChanged
        ch_stor.Checked = True
        ch_blance.Checked = True
        ch_items.Checked = True
        ch_addmonth.Checked = True
        ch_emp_add.Checked = True
        ch_emp_Show.Checked = True
        ch_emp_Search.Checked = True
        ch_emp_edit.Checked = True
        ch_emp_delete.Checked = True
        ch_emp_report.Checked = True
        ch_empatt_add.Checked = True
        ch_empatt_show.Checked = True
        ch_empatt_search.Checked = True
        ch_empatt_edit.Checked = True
        ch_empatt_delete.Checked = True
        ch_empatt_report.Checked = True
        ch_vend_add.Checked = True
        ch_vend_show.Checked = True
        ch_vend_search.Checked = True
        ch_vend_edit.Checked = True
        ch_vend_delete.Checked = True
        ch_vend_report.Checked = True
        ch_cust_add.Checked = True
        ch_cust_show.Checked = True
        ch_cust_search.Checked = True
        ch_cust_edit.Checked = True
        ch_cust_delete.Checked = True
        ch_cust_report.Checked = True
        ch_storvend_add.Checked = True
        ch_storvend_show.Checked = True
        ch_storvend_search.Checked = True
        ch_storvend_edit.Checked = True
        ch_storvend_delete.Checked = True
        ch_storvend_report.Checked = True
        ch_storcust_add.Checked = True
        ch_storcust_show.Checked = True
        ch_storcust_search.Checked = True
        ch_storcust_edit.Checked = True
        ch_storcust_delete.Checked = True
        ch_storcust_report.Checked = True
        ch_moneyimp_add.Checked = True
        ch_moneyimp_show.Checked = True
        ch_moneyimp_search.Checked = True
        ch_moneyimp_edit.Checked = True
        ch_moneyimp_delete.Checked = True
        ch_moneyimp_report.Checked = True
        ch_moneyexp_add.Checked = True
        ch_moneyexp_show.Checked = True
        ch_moneyexp_search.Checked = True
        ch_moneyexp_edit.Checked = True
        ch_moneyexp_delete.Checked = True
        ch_moneyexp_report.Checked = True

        ch_user_add.Checked = False
        ch_user_show.Checked = False
        ch_user_search.Checked = False
        ch_user_edit.Checked = False

    End Sub

    Private Sub rd_Employee_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rd_Employee.CheckedChanged
        ch_stor.Checked = True
        ch_blance.Checked = True
        ch_items.Checked = True
        ch_addmonth.Checked = True
        ch_emp_add.Checked = True
        ch_emp_Show.Checked = True
        ch_emp_Search.Checked = False
        ch_emp_edit.Checked = False
        ch_emp_delete.Checked = False
        ch_emp_report.Checked = True
        ch_empatt_add.Checked = True
        ch_empatt_show.Checked = True
        ch_empatt_search.Checked = False
        ch_empatt_edit.Checked = False
        ch_empatt_delete.Checked = False
        ch_empatt_report.Checked = True
        ch_vend_add.Checked = True
        ch_vend_show.Checked = True
        ch_vend_search.Checked = False
        ch_vend_edit.Checked = False
        ch_vend_delete.Checked = False
        ch_vend_report.Checked = True
        ch_cust_add.Checked = True
        ch_cust_show.Checked = True
        ch_cust_search.Checked = False
        ch_cust_edit.Checked = False
        ch_cust_delete.Checked = False
        ch_cust_report.Checked = True
        ch_storvend_add.Checked = True
        ch_storvend_show.Checked = True
        ch_storvend_search.Checked = False
        ch_storvend_edit.Checked = False
        ch_storvend_delete.Checked = False
        ch_storvend_report.Checked = True
        ch_storcust_add.Checked = True
        ch_storcust_show.Checked = True
        ch_storcust_search.Checked = False
        ch_storcust_edit.Checked = False
        ch_storcust_delete.Checked = False
        ch_storcust_report.Checked = True
        ch_moneyimp_add.Checked = True
        ch_moneyimp_show.Checked = True
        ch_moneyimp_search.Checked = False
        ch_moneyimp_edit.Checked = False
        ch_moneyimp_delete.Checked = False
        ch_moneyimp_report.Checked = True
        ch_moneyexp_add.Checked = True
        ch_moneyexp_show.Checked = True
        ch_moneyexp_search.Checked = False
        ch_moneyexp_edit.Checked = False
        ch_moneyexp_delete.Checked = False
        ch_moneyexp_report.Checked = True

        ch_user_add.Checked = False
        ch_user_show.Checked = False
        ch_user_search.Checked = False
        ch_user_edit.Checked = False

    End Sub
End Class