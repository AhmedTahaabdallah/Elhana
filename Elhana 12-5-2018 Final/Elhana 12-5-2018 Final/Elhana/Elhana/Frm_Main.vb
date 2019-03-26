Public Class Frm_Main

    Private Sub خروجToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles خروجToolStripMenuItem.Click
        Try
            cn1.Close()
            End
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            Dim g As String
            g = MsgBox("هل تريد الخروج من البرنامج ؟", MsgBoxStyle.YesNo, "تأكيد الخروج")

            If g = vbYes Then
                cn1.Close()
                End
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Frm_Emp.ShowDialog()
    End Sub

    Private Sub أضافةمستخدمToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles أضافةمستخدمToolStripMenuItem.Click
        Frm_Users.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Frm_EmpAttand.ShowDialog()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Frm_StoreVend.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Frm_Items.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Frm_Vendors.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Frm_Customers.ShowDialog()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Frm_Store.ShowDialog()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Frm_Money_Blance.ShowDialog()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Imports_Money.ShowDialog()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Export_Money.ShowDialog()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Frm_StoreCust.ShowDialog()
    End Sub

    Private Sub أضافةشهرToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles أضافةشهرToolStripMenuItem.Click
        frm_addmonth.ShowDialog()
    End Sub

    Private Sub Frm_Main_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        cn1.Close()
        End
    End Sub

    Private Sub Frm_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ut As Boolean
        ut = False
        auth_stor = rs_auth("User_stor").Value
        auth_blance = rs_auth("User_blance").Value
        auth_item = rs_auth("User_items").Value
        auth_addmonth = rs_auth("User_addmonth").Value

        If auth_stor = ut Then
            Button7.Visible = False
        Else
            Button7.Visible = True
        End If
        If auth_blance = ut Then
            Button13.Visible = False
        Else
            Button13.Visible = True
        End If
        If auth_item = ut Then
            Button5.Visible = False
        Else
            Button5.Visible = True
        End If
        If auth_addmonth = ut Then
            أضافةشهرToolStripMenuItem.Visible = False
        Else
            أضافةشهرToolStripMenuItem.Visible = True
        End If
        auth_useradd = rs_auth("User_user_add").Value
        auth_usershow = rs_auth("User_user_show").Value
        auth_usersearch = rs_auth("User_user_search").Value
        auth_useredit = rs_auth("User_user_edit").Value

        If auth_useradd = ut And auth_usershow = ut And auth_usersearch = ut And auth_useredit = ut Then
            أضافةمستخدمToolStripMenuItem.Visible = False
        Else
            أضافةمستخدمToolStripMenuItem.Visible = True
        End If

        auth_add = rs_auth("User_emp_add").Value
        auth_show = rs_auth("User_emp_show").Value
        auth_search = rs_auth("User_emp_search").Value
        auth_edit = rs_auth("User_emp_edit").Value
        auth_delete = rs_auth("User_emp_delete").Value
        auth_report = rs_auth("User_emp_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button1.Visible = False
        Else
            Button1.Visible = True
        End If
        auth_add = rs_auth("User_empatt_add").Value
        auth_show = rs_auth("User_empatt_show").Value
        auth_search = rs_auth("User_empatt_search").Value
        auth_edit = rs_auth("User_empatt_edit").Value
        auth_delete = rs_auth("User_empatt_delete").Value
        auth_report = rs_auth("User_empatt_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button4.Visible = False
        Else
            Button4.Visible = True
        End If
        auth_add = rs_auth("User_vend_add").Value
        auth_show = rs_auth("User_vend_show").Value
        auth_search = rs_auth("User_vend_search").Value
        auth_edit = rs_auth("User_vend_edit").Value
        auth_delete = rs_auth("User_vend_delete").Value
        auth_report = rs_auth("User_vend_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button2.Visible = False
        Else
            Button2.Visible = True
        End If
        auth_add = rs_auth("User_cust_add").Value
        auth_show = rs_auth("User_cust_show").Value
        auth_search = rs_auth("User_cust_search").Value
        auth_edit = rs_auth("User_cust_edit").Value
        auth_delete = rs_auth("User_cust_delete").Value
        auth_report = rs_auth("User_cust_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button3.Visible = False
        Else
            Button3.Visible = True
        End If
        auth_add = rs_auth("User_storvend_add").Value
        auth_show = rs_auth("User_storvend_show").Value
        auth_search = rs_auth("User_storvend_search").Value
        auth_edit = rs_auth("User_storvend_edit").Value
        auth_delete = rs_auth("User_storvend_delete").Value
        auth_report = rs_auth("User_storvend_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button10.Visible = False
        Else
            Button10.Visible = True
        End If
        auth_add = rs_auth("User_storcust_add").Value
        auth_show = rs_auth("User_storcust_show").Value
        auth_search = rs_auth("User_storcust_search").Value
        auth_edit = rs_auth("User_storcust_edit").Value
        auth_delete = rs_auth("User_storcust_delete").Value
        auth_report = rs_auth("User_storcust_reprt").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button9.Visible = False
        Else
            Button9.Visible = True
        End If
        auth_add = rs_auth("User_moneyimp_add").Value
        auth_show = rs_auth("User_moneyimp_show").Value
        auth_search = rs_auth("User_moneyimp_search").Value
        auth_edit = rs_auth("User_moneyimp_edit").Value
        auth_delete = rs_auth("User_moneyimp_delete").Value
        auth_report = rs_auth("User_moneyimp_repot").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button12.Visible = False
        Else
            Button12.Visible = True
        End If
        auth_add = rs_auth("User_moneyexp_add").Value
        auth_show = rs_auth("User_moneyexp_show").Value
        auth_search = rs_auth("User_moneyexp_search").Value
        auth_edit = rs_auth("User_moneyexp_edit").Value
        auth_delete = rs_auth("User_moneyexp_delete").Value
        auth_report = rs_auth("User_moneyexp_report").Value
        If auth_add = ut And auth_show = ut And auth_search = ut And auth_edit = ut And auth_delete = ut And auth_report = ut Then
            Button11.Visible = False
        Else
            Button11.Visible = True
        End If
        If u_Id = 1 Then
            Button7.Visible = True
            Button13.Visible = True
            Button5.Visible = True
            أضافةشهرToolStripMenuItem.Visible = True
            أضافةمستخدمToolStripMenuItem.Visible = True
            Button1.Visible = True
            Button4.Visible = True
            Button2.Visible = True
            Button3.Visible = True
            Button10.Visible = True
            Button9.Visible = True
            Button12.Visible = True
            Button11.Visible = True
        End If

    End Sub

    Private Sub تقفيلالشهورToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles تقفيلالشهورToolStripMenuItem.Click
        Me.Hide()
        frm_MonthesClosedLogin.Show()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        frm_custreview.Show()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        frm_backup.Show()
    End Sub
End Class