Imports ADODB
Public Class Frm_Login
    Public u_Name As Integer
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From dbo.Users Where User_Flag='" & s & "' And User_State_brok='" & s & "' And User_visbale='" & s & "'", cn1, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_SofUserName.EOF
                ComboBox1.Items.Add(rs_SofUserName("User_Name").Value)
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
            Dim u_name1 As String
            u_name1 = ComboBox1.SelectedItem.ToString()
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Users Where User_Name='" & u_name1 & "' And User_Flag='" & s & "' And User_State_brok='" & s & "' And User_visbale='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            u_Name = rs_SofUserName("User_ID").Value

            rs_SofUserName.Close()
            txt_Pass.Select()
           
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Frm_Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim appProc() As Process


            Dim strModName, strProcName As String
            strModName = Process.GetCurrentProcess.MainModule.ModuleName
            strProcName = System.IO.Path.GetFileNameWithoutExtension(strModName)


            appProc = Process.GetProcessesByName(strProcName)
            If appProc.Length > 1 Then
                MessageBox.Show("البرنامج قيد التشغيل", "تحذير")
                End
            Else

            End If
            'SQLEXPRESS
            con1 = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"

            gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"
            sql_str = "Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana"
            If cn1.State = 1 Then cn1.Close()
            cn1.Open(con1)
            cn12.Open("Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=UpCheackE;Data Source=PC-PC\SQLEXPRESS")
            Dim hj As String
            Dim hs As Boolean
            hs = False
            hj = "First"
            dy = Date.Now.Year
            dm = Date.Now.Month
            dd = Date.Now.Day
            dstri = dy + "-" + dm + "-" + dd
            If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
            rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
            rs_EmpAttand.Open("Select * From UpCheack Where datestr='" & dstri & "'", cn12, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then

            Else
                
                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                rs_EmpAttand.Open("Select * From UpCheack Where datestr='" & hj & "'", cn12, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_EmpAttand("state").Value = hs
                rs_EmpAttand.Update()

            End If

            If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
            rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
            rs_EmpAttand.Open("Select * From UpCheack Where datestr='" & hj & "'", cn12, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_EmpAttand("state").Value = False Then
                MessageBox.Show("برجاء المتابعة مع أحمد طه وذلك لتحديث البرنامج والتواصل معه على رقم  01141751527", "تحديث البرنامج")
                End
            End If

            'con1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Data Base\Tiba no2.accdb; Persist Security Info=False"
           
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Exit.Click
        Try
            cn1.Close()
            End
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_Enter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Enter.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If

        If txt_Pass.Text = "" Then
            MsgBox("من فضلك أدخل كلمة المرور", MsgBoxStyle.Information, "تنبيه")
            Exit Sub
        End If
        Dim u As String
        Dim p As String
        Dim f As Boolean
        Dim s As Boolean
        s = False
        Try
            u = u_Name
            p = txt_Pass.Text
            f = False

            If rs_auth.State = 1 Then rs_auth.Close()
            rs_auth.CursorLocation = CursorLocationEnum.adUseClient
            rs_auth.Open("Select * From Users Where User_ID=" & u & " And User_Password='" & p & "' And User_Flag='" & f & "' And User_State_brok='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_auth.EOF Or rs_auth.BOF Then
                MsgBox("أسم المستخدم أو كلمة السر غير صحيحة", MsgBoxStyle.Critical, "خطأ")
                Exit Sub
            End If
            If rs_auth("User_State").Value = f Then
                MsgBox("أسم المستخدم غير مفعل", MsgBoxStyle.Critical, "خطأ")
                Exit Sub
            End If
            u_Id = rs_auth("User_ID").Value
            user_name = rs_auth("User_Name").Value

            Me.Hide()
            Frm_Main.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try



    End Sub

    Private Sub txt_Pass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Pass.KeyPress
        If e.KeyChar = Convert.ToChar(13) Then
            If ComboBox1.Text = "" Then
                MsgBox("من فضلك أختر أسم المستخدم", MsgBoxStyle.Information, "تنبيه")
                Exit Sub
            End If

            If txt_Pass.Text = "" Then
                MsgBox("من فضلك أدخل كلمة المرور", MsgBoxStyle.Information, "تنبيه")
                Exit Sub
            End If
            Dim u As String
            Dim p As String
            Dim f As Boolean
            Dim s As Boolean
            s = False
            Try
                u = u_Name
                p = txt_Pass.Text
                f = False

                If rs_auth.State = 1 Then rs_auth.Close()
                rs_auth.CursorLocation = CursorLocationEnum.adUseClient
                rs_auth.Open("Select * From Users Where User_ID=" & u & " And User_Password='" & p & "' And User_Flag='" & f & "' And User_State_brok='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_auth.EOF Or rs_auth.BOF Then
                    MsgBox("أسم المستخدم أو كلمة السر غير صحيحة", MsgBoxStyle.Critical, "خطأ")
                    Exit Sub
                End If
                If rs_auth("User_State").Value = f Then
                    MsgBox("أسم المستخدم غير مفعل", MsgBoxStyle.Critical, "خطأ")
                    Exit Sub
                End If
                u_Id = rs_auth("User_ID").Value
                user_name = rs_auth("User_Name").Value


                Me.Hide()
                Frm_Main.Show()

            Catch ex As Exception
                MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            End Try


        End If
    End Sub

    
  
    
    Private Sub txt_Pass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Pass.TextChanged

    End Sub
End Class
