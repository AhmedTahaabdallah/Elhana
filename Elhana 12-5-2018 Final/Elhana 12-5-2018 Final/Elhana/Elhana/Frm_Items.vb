Imports ADODB
Public Class Frm_Items
    Public itm_id As Integer

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        TextBox1.Text = "0"
        TextBox3.Text = ""
        ComboBox1.Text = ""
        txt_Note.Text = ""
        ComboBox2.Enabled = False
        btn_All.Enabled = False
        btn_add.Enabled = False
        btn_save.Enabled = True
        TextBox3.Select()
    End Sub

    Private Sub Frm_Items_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        txt_Note.Text = ""
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Me.Dispose()
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل أسم الصنف", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            TextBox1.Text = 0
        End If
        If TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل سعر الصنف", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If

        Try
            Dim fg As Integer
            fg = 2
            Dim s As Boolean
            s = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Itemes Where Item_Name='" & TextBox3.Text & "' And Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From Itemes Where Item_ID= (SELECT MAX(Item_ID)  FROM Itemes)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("Item_ID").Value
                'u_name1 = u_name + 1

                If rs_Vendors.State = 1 Then rs_Vendors.Close()
                rs_Vendors.Open("Itemes", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Vendors.AddNew()
                'rs_Vendors("Item_ID").Value = u_name1
                rs_Vendors("Item_Name").Value = TextBox3.Text
                rs_Vendors("Item_DelFlag").Value = s
                rs_Vendors("User_ID").Value = u_Id
                rs_Vendors("Item_Price").Value = TextBox1.Text
                rs_Vendors("Item_Wight_Type").Value = ComboBox1.Text
                rs_Vendors("Item_Notes").Value = txt_Note.Text
                rs_Vendors("User_ID_edit").Value = fg
                rs_Vendors("User_ID_delete").Value = fg
                rs_Vendors.Update()
                MsgBox("تم حفظ الصنف بنجاح", MsgBoxStyle.Information, "حفظ بيانات الصنف")
                btn_All.Enabled = True
                btn_add.Enabled = True
                TextBox1.Text = ""
                TextBox3.Text = ""
                ComboBox1.Text = ""
                txt_Note.Text = ""
                btn_save.Enabled = False
                ComboBox2.Enabled = True
                rs_Vendors.Close()
            Else
                MsgBox("أسم الصنف موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                TextBox3.Text = ""
                TextBox3.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    
    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        TextBox1.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        txt_Note.Text = ""
        Frm_ShowItems.ShowDialog()
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Itemes Where Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_SofUserName.EOF
                ComboBox2.Items.Add(rs_SofUserName("Item_Name").Value)
                rs_SofUserName.MoveNext()
            Loop
            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim e_name1 As String
            e_name1 = ComboBox2.SelectedItem.ToString()
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Itemes Where Item_Name='" & e_name1 & "' And Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            TextBox3.Text = rs_SofUserName("Item_Name").Value
            txt_Note.Text = rs_SofUserName("Item_Notes").Value
            TextBox1.Text = rs_SofUserName("Item_Price").Value
            itm_id = rs_SofUserName("Item_ID").Value
            ComboBox1.Text = rs_SofUserName("Item_Wight_Type").Value
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.Open("SELECT * FROM Stor as st, Itemes as itm where st.Item_ID = itm.Item_ID And st.Item_ID =" & itm_id & " And itm.Item_ID=" & itm_id & " And st.Stor_DelFlag='" & s & "' And itm.Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_StoreCust.EOF Or rs_StoreCust.BOF Then
                ComboBox1.Enabled = True
            Else
                ComboBox1.Enabled = False
            End If

            'If IsDBNull(rs_StoreCust("c_No").Value) Then
            '    ComboBox1.Enabled = False
            'Else
            '    ComboBox1.Enabled = True
            'End If

            'rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        If TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل أسم الصنف", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            TextBox1.Text = 0
        End If
        If TextBox1.Text = "" Then
            MsgBox("من فضلك أدخل سعر الصنف", MsgBoxStyle.Information, "تنبيه")
            TextBox1.Select()
            Exit Sub
        End If
        If txt_Note.Text = "" Then
            txt_Note.Text = "لا يوجد ملاحظات"
        End If
        Try
            Dim g As String
            g = MsgBox("هل تريد تعديل هذا الصنف ؟", MsgBoxStyle.YesNo, "تأكيد تعديل")

            If g = vbYes Then
                Dim jj As Double
                jj = TextBox1.Text
                Dim s As Boolean
                s = False
                If rs_Emp.State = 1 Then rs_Emp.Close()
                rs_Emp.CursorLocation = CursorLocationEnum.adUseClient
                rs_Emp.Open("Select * From Itemes Where Item_ID=" & itm_id & " And Item_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Emp("Item_Name").Value = TextBox3.Text
                
                rs_Emp("Item_Price").Value = TextBox1.Text
                rs_Emp("Item_Wight_Type").Value = ComboBox1.Text
                rs_Emp("Item_Notes").Value = txt_Note.Text
                rs_Emp.Update()
                MsgBox("تم تعديل الصنف بنجاح", MsgBoxStyle.Information, "تعديل بيانات الصنف")
                 TextBox1.Text = ""
                TextBox3.Text = ""
                ComboBox1.Text = ""
                ComboBox2.Text = ""
                txt_Note.Text = ""
                rs_Emp.Close()
                ComboBox1.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class