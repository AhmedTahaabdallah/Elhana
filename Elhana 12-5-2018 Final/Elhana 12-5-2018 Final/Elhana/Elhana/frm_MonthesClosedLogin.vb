Imports ADODB
Public Class frm_MonthesClosedLogin
    Public u_Name As String
    Private Sub btn_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Exit.Click
        cn1.Close()
        End
    End Sub

    Private Sub frm_MonthesClosedLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From mUsers", cn1, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_SofUserName.EOF
                ComboBox1.Items.Add(rs_SofUserName("mUser_Name").Value)
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
            rs_SofUserName.Open("Select * From mUsers Where mUser_Name='" & u_name1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            u_Name = rs_SofUserName("mUser_Name").Value

            rs_SofUserName.Close()
            txt_Pass.Select()

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
        'Dim f As Boolean

        Try
            u = u_Name
            p = txt_Pass.Text
            'f = False

            If rs_U.State = 1 Then rs_U.Close()
            rs_U.CursorLocation = CursorLocationEnum.adUseClient
            rs_U.Open("Select * From mUsers Where mUser_Name='" & u & "' And mUser_Password='" & p & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_U.EOF Or rs_U.BOF Then
                MsgBox("أسم المستخدم أو كلمة السر غير صحيحة", MsgBoxStyle.Critical, "خطأ")
                Exit Sub
            End If
           
            monthscloseduser_Id = rs_U("mUser_ID").Value
            'user_name = rs_U("User_Name").Value
            rs_U.Close()
            Me.Hide()
            frm_MonthesClosed.Show()
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
            'Dim f As Boolean

            Try
                u = u_Name
                p = txt_Pass.Text
                'f = False

                If rs_U.State = 1 Then rs_U.Close()
                rs_U.CursorLocation = CursorLocationEnum.adUseClient
                rs_U.Open("Select * From mUsers Where mUser_Name='" & u & "' And mUser_Password='" & p & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_U.EOF Or rs_U.BOF Then
                    MsgBox("أسم المستخدم أو كلمة السر غير صحيحة", MsgBoxStyle.Critical, "خطأ")
                    Exit Sub
                End If

                monthscloseduser_Id = rs_U("mUser_ID").Value
                'user_name = rs_U("User_Name").Value
                rs_U.Close()
                Me.Hide()
                frm_MonthesClosed.Show()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "خطأ فى البرنامج")
            End Try
        End If
    End Sub

    
End Class