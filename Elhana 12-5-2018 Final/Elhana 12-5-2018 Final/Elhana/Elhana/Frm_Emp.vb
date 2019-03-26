Imports ADODB
Public Class Frm_Emp
    Public E_name As String
    Public E_id As String
    Private Sub Frm_Emp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_emp_add").Value
            auth_show = rs_auth("User_emp_show").Value
            auth_search = rs_auth("User_emp_search").Value
            auth_edit = rs_auth("User_emp_edit").Value
            auth_delete = rs_auth("User_emp_delete").Value
            auth_report = rs_auth("User_emp_report").Value
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

            TextBox1.Text = ""
            'TextBox2.Text = ""
            Txt_Salary.Text = ""
            txt_Name.Text = ""
            txt_Phone.Text = ""
            txt_Adrees.Text = ""
            txt_Age.Text = ""
            txt_Note.Text = ""
            CheckBox1.Checked = False
            TextBox3.Text = ""
            CheckBox3.Checked = False
            ComboBox1.Items.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_SofUserName.EOF
                ComboBox1.Items.Add(rs_SofUserName("Emp_Name").Value)
                rs_SofUserName.MoveNext()
            Loop
            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim e_name1 As String
            e_name1 = ComboBox1.SelectedItem.ToString()
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Name='" & e_name1 & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Txt_Salary.Text = rs_SofUserName("Emp_Salary").Value
            txt_Name.Text = rs_SofUserName("Emp_Name").Value
            E_name = rs_SofUserName("Emp_Name").Value
            E_id = rs_SofUserName("Emp_Id").Value
            txt_Phone.Text = rs_SofUserName("Emp_Phone").Value
            DateTimePicker1.Value = rs_SofUserName("Emp_Datework").Value
            txt_Adrees.Text = rs_SofUserName("adress").Value
            txt_Age.Text = rs_SofUserName("Emp_age").Value
            txt_Note.Text = rs_SofUserName("Emp_Note").Value
            CheckBox1.Checked = rs_SofUserName("Emp_State").Value
            TextBox1.Text = rs_SofUserName("Emp_Salary_Day").Value
            TextBox2.Text = rs_SofUserName("Emp_MonthCount").Value
            CheckBox2.Checked = rs_SofUserName("Emp_Salry_State").Value
            CheckBox3.Checked = rs_SofUserName("Emp_Shift_State").Value
            TextBox3.Text = rs_SofUserName("Emp_Shift_No").Value
            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Me.Dispose()
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        TextBox1.Text = ""
        'TextBox2.Text = ""
        txt_Name.Text = ""
        Txt_Salary.Text = ""
        txt_Phone.Text = "014"
        txt_Adrees.Text = "لا يوجد عنوان"
        txt_Age.Text = "20"
        txt_Note.Text = ""
        TextBox2.Text = ""
        CheckBox1.Checked = False
        ComboBox1.Items.Clear()
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        ComboBox1.Enabled = False
        CheckBox2.Checked = False
        btn_save.Enabled = True
        ComboBox1.Text = ""
        TextBox3.Text = ""
        CheckBox3.Checked = False
        txt_Name.Select()
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم الموظف", MsgBoxStyle.Information, "تنبيه")
            txt_Name.Select()
            Exit Sub
        End If

        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            txt_Phone.Select()
            Exit Sub
        End If
        If txt_Adrees.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            txt_Adrees.Select()
            Exit Sub
        End If
        If txt_Age.Text = "" Then
            MsgBox("من فضلك أدخل سن الموظف", MsgBoxStyle.Information, "تنبيه")

            txt_Age.Select()
            Exit Sub
        End If
        If Txt_Salary.Text = "" Then
            MsgBox("من فضلك أدخل الراتب", MsgBoxStyle.Information, "تنبيه")

            Txt_Salary.Select()
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MsgBox("من فضلك أدخل عدد أيام الشهر", MsgBoxStyle.Information, "تنبيه")

            TextBox2.Select()
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل عدد الراتب اليومى", MsgBoxStyle.Information, "تنبيه")

            TextBox1.Select()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل عدد ساعات العمل اليومى", MsgBoxStyle.Information, "تنبيه")

            TextBox3.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Emp.State = 1 Then rs_Emp.Close()
            rs_Emp.CursorLocation = CursorLocationEnum.adUseClient
            rs_Emp.Open("Select * From Employyes Where Emp_Name='" & txt_Name.Text & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Emp.EOF Or rs_Emp.BOF Then
                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From Employyes Where Emp_Id= (SELECT MAX(Emp_Id)  FROM Employyes)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("Emp_Id").Value
                'u_name1 = u_name + 1

                Dim uk As Integer
                uk = 2
                Dim jj As Double
                jj = TextBox1.Text
                If rs_Emp.State = 1 Then rs_Emp.Close()
                rs_Emp.Open("Employyes", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Emp.AddNew()
                'rs_Emp("Emp_Id").Value = txt_Name.Text
                rs_Emp("Emp_Name").Value = txt_Name.Text
                rs_Emp("Emp_Phone").Value = txt_Phone.Text
                rs_Emp("Emp_Datework").Value = DateTimePicker1.Value.ToShortDateString()
                rs_Emp("Emp_Flag").Value = s
                rs_Emp("User_ID").Value = u_Id
                rs_Emp("adress").Value = txt_Adrees.Text
                rs_Emp("Emp_Salary").Value = Txt_Salary.Text
                rs_Emp("Emp_age").Value = txt_Age.Text
                rs_Emp("Emp_Note").Value = txt_Note.Text
                rs_Emp("Emp_State").Value = CheckBox1.Checked
                rs_Emp("Emp_Salry_State").Value = CheckBox2.Checked
                rs_Emp("Emp_Salary_Day").Value = jj
                rs_Emp("Emp_MonthCount").Value = TextBox2.Text
                rs_Emp("Emp_Shift_State").Value = CheckBox3.Checked
                rs_Emp("Emp_Shift_No").Value = TextBox3.Text
                rs_Emp("User_ID_edit").Value = uk
                rs_Emp("User_ID_delete").Value = uk
                rs_Emp.Update()
                MsgBox("تم حفظ الموظف بنجاح", MsgBoxStyle.Information, "حفظ بيانات الموظف")
                TextBox1.Text = ""
                'TextBox2.Text = ""
                txt_Name.Text = ""
                Txt_Salary.Text = ""
                txt_Phone.Text = ""
                TextBox2.Text = ""
                txt_Adrees.Text = ""
                txt_Age.Text = ""
                txt_Note.Text = ""
                ComboBox1.Items.Clear()
                btn_All.Enabled = True
                btn_etite.Enabled = True
                btn_delete.Enabled = True
                btn_add.Enabled = True
                ComboBox1.Enabled = True
                btn_save.Enabled = False
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                ComboBox1.Text = ""
                TextBox3.Text = ""
                CheckBox3.Checked = False
                rs_Emp.Close()
            Else
                MsgBox("أسم الموظف موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                txt_Name.Text = ""
                txt_Name.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك اختر أسم الموظف", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If

        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم الموظف", MsgBoxStyle.Information, "تنبيه")
            txt_Name.Select()
            Exit Sub
        End If

        If txt_Phone.Text = "" Then
            MsgBox("من فضلك أدخل رقم الموبايل", MsgBoxStyle.Information, "تنبيه")
            txt_Phone.Select()
            Exit Sub
        End If
        If txt_Adrees.Text = "" Then
            MsgBox("من فضلك أدخل العنوان", MsgBoxStyle.Information, "تنبيه")
            txt_Adrees.Select()
            Exit Sub
        End If
        If txt_Age.Text = "" Then
            MsgBox("من فضلك أدخل سن الموظف", MsgBoxStyle.Information, "تنبيه")
            txt_Age.Select()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل عدد ساعات العمل اليومى", MsgBoxStyle.Information, "تنبيه")

            TextBox3.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim g As String
            g = MsgBox("هل تريد تعديل هذا الموظف ؟", MsgBoxStyle.YesNo, "تأكيد تعديل")

            If g = vbYes Then
                Dim jj As Double
                jj = TextBox1.Text
                Dim s As Boolean
                s = False
                If rs_Emp.State = 1 Then rs_Emp.Close()
                rs_Emp.CursorLocation = CursorLocationEnum.adUseClient
                rs_Emp.Open("Select * From Employyes Where Emp_Name='" & E_name & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Emp("Emp_Name").Value = txt_Name.Text
                rs_Emp("Emp_Phone").Value = txt_Phone.Text
                rs_Emp("Emp_Datework").Value = DateTimePicker1.Value.ToShortDateString()
                rs_Emp("User_ID_edit").Value = u_Id
                rs_Emp("adress").Value = txt_Adrees.Text
                rs_Emp("Emp_age").Value = txt_Age.Text
                rs_Emp("Emp_Note").Value = txt_Note.Text
                rs_Emp("Emp_State").Value = CheckBox1.Checked
                rs_Emp("Emp_Salry_State").Value = CheckBox2.Checked
                rs_Emp("Emp_Salary").Value = Txt_Salary.Text
                rs_Emp("Emp_Salary_Day").Value = jj
                rs_Emp("Emp_MonthCount").Value = TextBox2.Text
                rs_Emp("Emp_Shift_State").Value = CheckBox3.Checked
                rs_Emp("Emp_Shift_No").Value = TextBox3.Text
                rs_Emp.Update()
                MsgBox("تم تعديل الموظف بنجاح", MsgBoxStyle.Information, "تعديل بيانات الموظف")
                TextBox1.Text = ""
                'TextBox2.Text = ""
                txt_Name.Text = ""
                Txt_Salary.Text = ""
                txt_Phone.Text = ""
                txt_Adrees.Text = ""
                txt_Age.Text = ""
                txt_Note.Text = ""
                TextBox2.Text = ""
                ComboBox1.Items.Clear()
                ComboBox1.Text = ""
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                TextBox3.Text = ""
                CheckBox3.Checked = False
                rs_Emp.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم الموظف", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Employyes_Presence Where Emp_Id='" & E_id & "' And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
               

                Dim g As String
                g = MsgBox("هل تريد حذف هذا الموظف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then
                    
                    If rs_Emp.State = 1 Then rs_Emp.Close()
                    rs_Emp.CursorLocation = CursorLocationEnum.adUseClient
                    rs_Emp.Open("Select * From Employyes Where Emp_Id='" & E_id & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                    s = True
                    rs_Emp("Emp_Flag").Value = s
                    rs_Emp("User_ID_delete").Value = u_Id
                    rs_Emp.Update()
                    MsgBox("تم حذف الموظف بنجاح", MsgBoxStyle.Information, "حذف بيانات الموظف")
                    TextBox1.Text = ""
                    'TextBox2.Text = ""
                    txt_Name.Text = ""
                    Txt_Salary.Text = ""
                    txt_Phone.Text = ""
                    txt_Adrees.Text = ""
                    TextBox2.Text = ""
                    txt_Age.Text = ""
                    txt_Note.Text = ""
                    ComboBox1.Items.Clear()
                    ComboBox1.Text = ""
                    CheckBox1.Checked = False
                    CheckBox2.Checked = False
                    TextBox3.Text = ""
                    CheckBox3.Checked = False
                    rs_Emp.Close()
                End If
            Else
                MsgBox("لا يمكن حذف هذا الموظف لوجود أذونات حضور وأنصراف لهذا الموظف", MsgBoxStyle.Information, "تحذير")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

   

    
    Private Sub txt_Phone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Phone.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub Txt_Salary_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txt_Salary.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub Txt_Salary_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txt_Salary.TextChanged
        Try
            Dim sal As Integer
            Dim mo As Integer

            sal = Val(Txt_Salary.Text)
            mo = Val(TextBox2.Text)

            TextBox1.Text = sal / mo
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
        Try
            Dim sal As Integer
            Dim mo As Integer

            sal = Val(Txt_Salary.Text)
            mo = Val(TextBox2.Text)
            TextBox1.Text = sal / mo
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_Phone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Phone.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Frm_UpdateCount_Month.ShowDialog()
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Try
            If CheckBox3.Checked = True Then
                TextBox3.Enabled = True
                TextBox3.Text = 8
            End If
            If CheckBox3.Checked = False Then
                TextBox3.Enabled = False
                TextBox3.Text = 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        Frm_ShowEmp.ShowDialog()
    End Sub
End Class