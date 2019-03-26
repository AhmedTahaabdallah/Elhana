Imports ADODB
Imports System.Data.OleDb
Public Class frm_MonthesClosed
    Public month_idnew As Integer
    Public month_idold As Integer
    Public month_idsearch As Integer
   

   

    Private Sub frm_MonthesClosed_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        cn1.Close()
        End
    End Sub

    Private Sub frm_MonthesClosed_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'ComboBox2.Text = ""
        'txt_Finaltotalblance.Text = ""
        'txt_lasttotalblance.Text = ""
        'txt_Note.Text = ""
        'txt_totalAhmed.Text = ""
        'txt_totalallcustmoney.Text = ""
        'txt_totalAllquntityprice.Text = ""
        'txt_totalbeforefinal.Text = ""
        'txt_totalbenfet.Text = ""
        'txt_totalblance.Text = ""
        'txt_TotalBlasticPrice.Text = ""
        'txt_totalEslam.Text = ""
        'txt_totalvendblance.Text = ""
        'ComboBox3.Text = ""
        'TextBox1.Text = ""
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

   

    Private Sub txt_totalAllquntityprice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalAllquntityprice.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalAllquntityprice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalAllquntityprice.TextChanged

        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalblance_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalblance.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalblance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalblance.TextChanged
        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalallcustmoney_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalallcustmoney.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalallcustmoney_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalallcustmoney.TextChanged
        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_TotalBlasticPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_TotalBlasticPrice.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_TotalBlasticPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_TotalBlasticPrice.TextChanged
        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalEslam_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalEslam.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalEslam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalEslam.TextChanged
        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalAhmed_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalAhmed.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalAhmed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalAhmed.TextChanged
        'If ComboBox2.Text = "" Then
        '    MsgBox("أدخل أسم الشهر الحالى الأول", MsgBoxStyle.Information, "تحذير")
        '    ComboBox2.Select()
        '    Exit Sub
        'End If
        Try
            Dim taqp As Integer
            Dim tb As Integer
            Dim tacm As Integer
            Dim tBP As Integer
            Dim tE As Integer
            Dim tA As Integer
            taqp = Val(txt_totalAllquntityprice.Text)
            tb = Val(txt_totalblance.Text)
            tacm = Val(txt_totalallcustmoney.Text)
            tBP = Val(txt_TotalBlasticPrice.Text)
            tE = Val(txt_totalEslam.Text)
            tA = Val(txt_totalAhmed.Text)
            txt_totalbeforefinal.Text = taqp + tb + tacm + tBP + tE + tA
            TextBox1.Text = Val(txt_totalbeforefinal.Text) - Val(txt_totalvendblance.Text)
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalbeforefinal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalbeforefinal.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalbeforefinal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalbeforefinal.TextChanged
        Try
            Dim tef As Integer
            Dim tvb As Integer
            tef = Val(txt_totalbeforefinal.Text)
            tvb = Val(txt_totalvendblance.Text)
            TextBox1.Text = tef - tvb
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub txt_totalvendblance_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_totalvendblance.KeyPress
        If Asc(e.KeyChar) <> 8 Then

            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If

        End If
    End Sub

    Private Sub txt_totalvendblance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_totalvendblance.TextChanged
        Try
            Dim tef As Integer
            Dim tvb As Integer
            tef = Val(txt_totalbeforefinal.Text)
            tvb = Val(txt_totalvendblance.Text)
            TextBox1.Text = tef - tvb
            txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
            txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox2.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
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
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_idnew = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub



    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        ComboBox2.Text = ""
        txt_Finaltotalblance.Text = ""
        txt_lasttotalblance.Text = ""
        txt_Note.Text = ""
        txt_totalAhmed.Text = ""
        txt_totalallcustmoney.Text = ""
        txt_totalAllquntityprice.Text = ""
        txt_totalbeforefinal.Text = ""
        txt_totalbenfet.Text = ""
        txt_totalblance.Text = ""
        txt_TotalBlasticPrice.Text = ""
        txt_totalEslam.Text = ""
        txt_totalvendblance.Text = ""
        ComboBox3.Text = ""
        TextBox1.Text = ""
        btn_Add.Enabled = False
        btn_save.Enabled = True
        btnedit.Enabled = False
        ComboBox2.Enabled = True
        btn_All.Enabled = False
        dtn_delete.Enabled = False
        GroupBox3.Enabled = False
        ComboBox2.Select()
    End Sub

    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox3.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox3.Items.Add(rs_Cars("Month_Name").Value)
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
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Dim rrr As Integer
            rrr = rs_Cars("Month_ID").Value
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.Open("Select * From month_closed Where Month_ID=" & rrr & " And mc_DeleteFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Cars.EOF Or rs_Cars.BOF Then
                MsgBox("هذا الشهر لم يتم عمل تقفيل له من قبل أختر شهر أخر", MsgBoxStyle.Information, "تنبيه")
                ComboBox3.Select()

            Else
                txt_lasttotalblance.Text = rs_Cars("mc_lasttotalblance").Value
                txt_totalbenfet.Text = Val(TextBox1.Text) - Val(txt_lasttotalblance.Text)
                txt_Finaltotalblance.Text = Val(TextBox1.Text) - Val(txt_totalbenfet.Text)
                month_idold = rrr
            End If

            rs_Cars.Close()
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
            Dim mmmm As Integer
            mmmm = rs_Cars("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.Open("Select * From month_closed Where Month_ID=" & mmmm & " And mc_DeleteFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_U.EOF Or rs_U.BOF Then
                MsgBox("هذا الشهر لم يتم عمل تقفيل له من قبل أختر شهر أخر", MsgBoxStyle.Information, "تنبيه")
                ComboBox1.Select()
                rs_U.Close()
            Else

                txt_totalAllquntityprice.Text = rs_U("mc_totalAllquntityprice").Value
                txt_totalblance.Text = rs_U("mc_totalblance").Value
                txt_totalallcustmoney.Text = rs_U("mc_totalallcustmoney").Value
                txt_TotalBlasticPrice.Text = rs_U("mc_TotalBlasticPrice").Value
                txt_totalEslam.Text = rs_U("mc_totalEslam").Value
                txt_totalAhmed.Text = rs_U("mc_totalAhmed").Value
                txt_totalbeforefinal.Text = rs_U("mc_totalbeforefinal").Value
                txt_totalvendblance.Text = rs_U("mc_totalvendblance").Value
                TextBox1.Text = rs_U("mc_firsttotal").Value
                txt_lasttotalblance.Text = rs_U("mc_lasttotalblance").Value
                txt_totalbenfet.Text = rs_U("mc_totalbenfet").Value
                txt_Finaltotalblance.Text = rs_U("mc_Finaltotalblance").Value
                txt_Note.Text = rs_U("mc_notes").Value
                Dim hj As Integer
                hj = rs_U("Month_ID").Value
                s = False
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Select * From year_monthes Where Month_ID=" & hj & " And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Customers.EOF Or rs_Customers.BOF Then
                Else
                    ComboBox2.Text = rs_Customers("Month_Name").Value
                End If
                Dim hjr As Integer
                hjr = rs_U("Month_IDlast").Value
                month_idold = rs_U("Month_IDlast").Value
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Select * From year_monthes Where Month_ID=" & hjr & " And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If rs_Customers.EOF Or rs_Customers.BOF Then
                Else
                    ComboBox3.Text = rs_Customers("Month_Name").Value

                End If
                ComboBox2.Enabled = False
                month_idsearch = mmmm
                rs_U.Close()

                rs_Customers.Close()
            End If

            'month_idsearch = rs_Cars("Month_ID").Value
            'rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If
       
        
        If txt_totalAllquntityprice.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر البضاعة", MsgBoxStyle.Information, "تنبيه")
            txt_totalAllquntityprice.Select()
            Exit Sub
        End If
        If txt_totalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر الخزنة", MsgBoxStyle.Information, "تنبيه")
            txt_totalblance.Select()
            Exit Sub
        End If
        If txt_totalallcustmoney.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى الأجل", MsgBoxStyle.Information, "تنبيه")
            txt_totalallcustmoney.Select()
            Exit Sub
        End If
        If txt_TotalBlasticPrice.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر البلاستيك", MsgBoxStyle.Information, "تنبيه")
            txt_TotalBlasticPrice.Select()
            Exit Sub
        End If
       
        If txt_totalEslam.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى مسحوبات أسلام", MsgBoxStyle.Information, "تنبيه")
            txt_totalEslam.Select()
            Exit Sub
        End If
        If txt_totalAhmed.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى مسحوبات أحمد", MsgBoxStyle.Information, "تنبيه")
            txt_totalAhmed.Select()
            Exit Sub
        End If
        If txt_totalbeforefinal.Text = "" Then
            MsgBox("من فضلك أدخل الاجمالى ", MsgBoxStyle.Information, "تنبيه")
            txt_totalbeforefinal.Select()
            Exit Sub
        End If
        If txt_totalvendblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى اللى علينا", MsgBoxStyle.Information, "تنبيه")
            txt_totalvendblance.Select()
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل الاجمالى", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If
        
        If ComboBox3.Text = "" Then
            MsgBox("من فضلك أخنر الشهر السابق", MsgBoxStyle.Information, "تنبيه")
            ComboBox3.Select()
            Exit Sub
        End If
        If txt_lasttotalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى الرصيد", MsgBoxStyle.Information, "تنبيه")
            txt_lasttotalblance.Select()
            Exit Sub
        End If
        If txt_totalbenfet.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى ربح الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            txt_totalbenfet.Select()
            Exit Sub
        End If
        If txt_Finaltotalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى رصيد الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            txt_Finaltotalblance.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If

        Try
            Dim s As Boolean
            s = False
            Dim uk As Integer
            uk = 2
            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Select * From month_closed Where Month_ID=" & month_idnew & " And mc_DeleteFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Customers.EOF Or rs_Customers.BOF Then
               
                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("month_closed", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Customers.AddNew()
                rs_Customers("Month_ID").Value = month_idnew
                rs_Customers("mc_totalAllquntityprice").Value = txt_totalAllquntityprice.Text
                rs_Customers("mc_totalblance").Value = txt_totalblance.Text
                rs_Customers("mc_totalallcustmoney").Value = txt_totalallcustmoney.Text
                rs_Customers("mc_TotalBlasticPrice").Value = txt_TotalBlasticPrice.Text
                rs_Customers("mc_totalEslam").Value = txt_totalEslam.Text
                rs_Customers("mc_totalAhmed").Value = txt_totalAhmed.Text
                rs_Customers("mc_totalbeforefinal").Value = txt_totalbeforefinal.Text
                rs_Customers("mc_totalvendblance").Value = txt_totalvendblance.Text
                rs_Customers("mc_firsttotal").Value = TextBox1.Text
                rs_Customers("Month_IDlast").Value = month_idold
                rs_Customers("mc_lasttotalblance").Value = txt_lasttotalblance.Text
                rs_Customers("mc_totalbenfet").Value = txt_totalbenfet.Text
                rs_Customers("mc_Finaltotalblance").Value = txt_Finaltotalblance.Text
                rs_Customers("mc_notes").Value = txt_Note.Text
                rs_Customers("mc_DeleteFlag").Value = s
                rs_Customers("User_ID").Value = u_Id
                rs_Customers("User_ID_edit").Value = uk
                rs_Customers("User_ID_delete").Value = uk
                rs_Customers("monthscloseduser_Id").Value = monthscloseduser_Id
                rs_Customers.Update()
                MsgBox("تم حفظ الشهر بنجاح", MsgBoxStyle.Information, "حفظ بيانات الشهر")
                ComboBox2.Text = ""
                txt_Finaltotalblance.Text = ""
                txt_lasttotalblance.Text = ""
                txt_Note.Text = ""
                txt_totalAhmed.Text = ""
                txt_totalallcustmoney.Text = ""
                txt_totalAllquntityprice.Text = ""
                txt_totalbeforefinal.Text = ""
                txt_totalbenfet.Text = ""
                txt_totalblance.Text = ""
                txt_TotalBlasticPrice.Text = ""
                txt_totalEslam.Text = ""
                txt_totalvendblance.Text = ""
                ComboBox3.Text = ""
                TextBox1.Text = ""
                btn_Add.Enabled = True
                btn_save.Enabled = False
                btnedit.Enabled = True
                btn_All.Enabled = True
                dtn_delete.Enabled = True
                GroupBox3.Enabled = True
            Else
                MsgBox("الشهر موجود من قبل أختر شهر أخر", MsgBoxStyle.Information, "تحذير")
                ComboBox2.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btnedit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnedit.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر ", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MsgBox("من فضلك أختر الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            ComboBox2.Select()
            Exit Sub
        End If

        If txt_totalAllquntityprice.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر البضاعة", MsgBoxStyle.Information, "تنبيه")
            txt_totalAllquntityprice.Select()
            Exit Sub
        End If
        If txt_totalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر الخزنة", MsgBoxStyle.Information, "تنبيه")
            txt_totalblance.Select()
            Exit Sub
        End If
        If txt_totalallcustmoney.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى الأجل", MsgBoxStyle.Information, "تنبيه")
            txt_totalallcustmoney.Select()
            Exit Sub
        End If
        If txt_TotalBlasticPrice.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى سعر البلاستيك", MsgBoxStyle.Information, "تنبيه")
            txt_TotalBlasticPrice.Select()
            Exit Sub
        End If

        If txt_totalEslam.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى مسحوبات أسلام", MsgBoxStyle.Information, "تنبيه")
            txt_totalEslam.Select()
            Exit Sub
        End If
        If txt_totalAhmed.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى مسحوبات أحمد", MsgBoxStyle.Information, "تنبيه")
            txt_totalAhmed.Select()
            Exit Sub
        End If
        If txt_totalbeforefinal.Text = "" Then
            MsgBox("من فضلك أدخل الاجمالى ", MsgBoxStyle.Information, "تنبيه")
            txt_totalbeforefinal.Select()
            Exit Sub
        End If
        If txt_totalvendblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى اللى علينا", MsgBoxStyle.Information, "تنبيه")
            txt_totalvendblance.Select()
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل الاجمالى", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If

        If ComboBox3.Text = "" Then
            MsgBox("من فضلك أخنر الشهر السابق", MsgBoxStyle.Information, "تنبيه")
            ComboBox3.Select()
            Exit Sub
        End If
        If txt_lasttotalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى الرصيد", MsgBoxStyle.Information, "تنبيه")
            txt_lasttotalblance.Select()
            Exit Sub
        End If
        If txt_totalbenfet.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى ربح الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            txt_totalbenfet.Select()
            Exit Sub
        End If
        If txt_Finaltotalblance.Text = "" Then
            MsgBox("من فضلك أدخل إجمالى رصيد الشهر الحالى", MsgBoxStyle.Information, "تنبيه")
            txt_Finaltotalblance.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim s As Boolean
            s = False
            If rs_Customers.State = 1 Then rs_Customers.Close()
            rs_Customers.Open("Select * From month_closed Where Month_ID=" & month_idsearch & " And mc_DeleteFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_Customers.EOF Or rs_Customers.BOF Then

            Else
                'rs_Customers("Month_ID").Value = month_idnew
                rs_Customers("mc_totalAllquntityprice").Value = txt_totalAllquntityprice.Text
                rs_Customers("mc_totalblance").Value = txt_totalblance.Text
                rs_Customers("mc_totalallcustmoney").Value = txt_totalallcustmoney.Text
                rs_Customers("mc_TotalBlasticPrice").Value = txt_TotalBlasticPrice.Text
                rs_Customers("mc_totalEslam").Value = txt_totalEslam.Text
                rs_Customers("mc_totalAhmed").Value = txt_totalAhmed.Text
                rs_Customers("mc_totalbeforefinal").Value = txt_totalbeforefinal.Text
                rs_Customers("mc_totalvendblance").Value = txt_totalvendblance.Text
                rs_Customers("mc_firsttotal").Value = TextBox1.Text
                rs_Customers("Month_IDlast").Value = month_idold
                rs_Customers("mc_lasttotalblance").Value = txt_lasttotalblance.Text
                rs_Customers("mc_totalbenfet").Value = txt_totalbenfet.Text
                rs_Customers("mc_Finaltotalblance").Value = txt_Finaltotalblance.Text
                rs_Customers("mc_notes").Value = txt_Note.Text
                rs_Customers("mc_DeleteFlag").Value = s
                rs_Customers("User_ID_edit").Value = u_Id
                rs_Customers("monthscloseduser_Id").Value = monthscloseduser_Id
                rs_Customers.Update()
                MsgBox("تم تعديل الشهر بنجاح", MsgBoxStyle.Information, "تعديل بيانات الشهر")
                ComboBox2.Text = ""
                txt_Finaltotalblance.Text = ""
                txt_lasttotalblance.Text = ""
                txt_Note.Text = ""
                txt_totalAhmed.Text = ""
                txt_totalallcustmoney.Text = ""
                txt_totalAllquntityprice.Text = ""
                txt_totalbeforefinal.Text = ""
                txt_totalbenfet.Text = ""
                txt_totalblance.Text = ""
                txt_TotalBlasticPrice.Text = ""
                txt_totalEslam.Text = ""
                txt_totalvendblance.Text = ""
                ComboBox3.Text = ""
                TextBox1.Text = ""
                btn_Add.Enabled = True
                btn_save.Enabled = False
                btnedit.Enabled = True
                btn_All.Enabled = True
                GroupBox3.Enabled = True
                ComboBox2.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub dtn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtn_delete.Click
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر أسم الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        Try
            Dim s As Boolean
            s = False
            
                Dim g As String
            g = MsgBox("هل تريد حذف هذا الشهر ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then

                If rs_Customers.State = 1 Then rs_Customers.Close()
                rs_Customers.Open("Select * From month_closed Where Month_ID='" & month_idsearch & "' And mc_DeleteFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                s = True
                rs_Customers("mc_DeleteFlag").Value = s
                rs_Customers("User_ID_delete").Value = u_Id
                rs_Customers.Update()
                MsgBox("تم حذف الشهر بنجاح", MsgBoxStyle.Information, "حذف بيانات الشهر")
                ComboBox2.Text = ""
                txt_Finaltotalblance.Text = ""
                txt_lasttotalblance.Text = ""
                txt_Note.Text = ""
                txt_totalAhmed.Text = ""
                txt_totalallcustmoney.Text = ""
                txt_totalAllquntityprice.Text = ""
                txt_totalbeforefinal.Text = ""
                txt_totalbenfet.Text = ""
                txt_totalblance.Text = ""
                txt_TotalBlasticPrice.Text = ""
                txt_totalEslam.Text = ""
                txt_totalvendblance.Text = ""
                ComboBox3.Text = ""
                TextBox1.Text = ""
                btn_Add.Enabled = True
                btn_save.Enabled = False
                btnedit.Enabled = True
                btn_All.Enabled = True
                ComboBox2.Enabled = True
                GroupBox3.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        frm_Showmonthsclosed.ShowDialog()
    End Sub
End Class