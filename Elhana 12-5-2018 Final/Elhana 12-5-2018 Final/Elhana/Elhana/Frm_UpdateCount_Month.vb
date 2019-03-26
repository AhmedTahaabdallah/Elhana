Public Class Frm_UpdateCount_Month

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            If TextBox3.Text = "" Then
                MsgBox("من فضلك أدخل عدد أيام الشهر", MsgBoxStyle.Information, "تنبيه")
                TextBox3.Select()
                Exit Sub
            End If

            If TextBox3.Text = 30 Or TextBox3.Text = 31 Or TextBox3.Text = 29 Or TextBox3.Text = 28 Then
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                Dim s As Boolean
                s = False
                Dim s1 As Boolean
                s1 = True
                rs_SofUserName.Open("Select * From Employyes", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim d1 As Integer
                Dim dold As Integer
                Dim dnew As Double
                d1 = TextBox3.Text
                Do While Not rs_SofUserName.EOF
                    dold = rs_SofUserName("Emp_Salary").Value
                    dnew = Val(dold) / Val(d1)
                    rs_SofUserName("Emp_MonthCount").Value = d1
                    rs_SofUserName("Emp_Salary_Day").Value = dnew
                    rs_SofUserName.Update()
                    rs_SofUserName.MoveNext()
                Loop
                rs_SofUserName.Close()
                Me.Dispose()
            Else
                MsgBox("من فضلك أدخل عدد أيام الشهر صحيحة", MsgBoxStyle.Information, "تنبيه")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    
End Class